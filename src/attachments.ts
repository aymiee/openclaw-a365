import { randomUUID } from "node:crypto";
import { mkdir, readdir, rm, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

export type AttachmentsLog = {
  info: (msg: string, meta?: unknown) => void;
  debug?: (msg: string, meta?: unknown) => void;
  warn?: (msg: string, meta?: unknown) => void;
  error: (msg: string, meta?: unknown) => void;
};

export type DownloadedAttachments = {
  mediaPaths: string[];
  mediaTypes: string[];
  mediaDir?: string;
};

type Attachment = {
  contentType?: string;
  contentUrl?: string;
  content?: unknown;
  name?: string;
};

type DownloadedFile = {
  content: Buffer;
  contentType: string;
};

const CACHE_ROOT = join(tmpdir(), "openclaw-a365");
const TTL_MS = 60 * 60 * 1000;

const EXT_BY_CONTENT_TYPE: Record<string, string> = {
  "image/png": "png",
  "image/jpeg": "jpg",
  "image/jpg": "jpg",
  "image/gif": "gif",
  "image/webp": "webp",
  "image/bmp": "bmp",
  "image/svg+xml": "svg",
  "application/pdf": "pdf",
  "text/plain": "txt",
  "text/csv": "csv",
  "application/json": "json",
  "application/msword": "doc",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document": "docx",
  "application/vnd.ms-excel": "xls",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "xlsx",
  "application/vnd.ms-powerpoint": "ppt",
  "application/vnd.openxmlformats-officedocument.presentationml.presentation": "pptx",
};

function sanitizePathSegment(input: string): string {
  return input.replace(/[^A-Za-z0-9._-]+/g, "_").slice(0, 120) || "_";
}

function pickFilename(index: number, originalName: string | undefined, contentType: string): string {
  if (originalName) {
    const safe = sanitizePathSegment(originalName);
    if (safe && safe !== "_") {
      return safe;
    }
  }
  const ext = EXT_BY_CONTENT_TYPE[contentType.toLowerCase()] ?? "bin";
  return `attachment-${index}.${ext}`;
}

function encodeShareUrl(url: string): string {
  const base64 = Buffer.from(url, "utf-8").toString("base64");
  return `u!${base64.replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "")}`;
}

async function downloadSingleAttachment(
  att: Attachment,
  graphToken: string | undefined,
  log: AttachmentsLog,
): Promise<DownloadedFile | undefined> {
  if (!att.contentType) return undefined;
  if (att.contentType.startsWith("text/html")) return undefined;
  if (att.contentType.startsWith("application/vnd.microsoft.card.")) return undefined;

  let contentObj: Record<string, unknown> | undefined;
  if (typeof att.content === "string") {
    try {
      contentObj = JSON.parse(att.content);
    } catch {
      contentObj = undefined;
    }
  } else if (att.content && typeof att.content === "object") {
    contentObj = att.content as Record<string, unknown>;
  }

  const downloadUrl =
    typeof contentObj?.downloadUrl === "string" ? contentObj.downloadUrl : undefined;

  log.info("attachment received", {
    type: att.contentType,
    name: att.name ?? "(none)",
    hasDownloadUrl: Boolean(downloadUrl),
    hasContentUrl: Boolean(att.contentUrl),
  });

  if (downloadUrl) {
    try {
      const resp = await fetch(downloadUrl);
      if (resp.ok) {
        const buf = Buffer.from(await resp.arrayBuffer());
        return {
          content: buf,
          contentType: resp.headers.get("content-type") ?? att.contentType,
        };
      }
      log.error(`downloadUrl fetch failed: ${resp.status} for ${downloadUrl.slice(0, 120)}`);
    } catch (err) {
      log.error(`downloadUrl fetch error: ${String(err)}`);
    }
  }

  if (att.contentUrl && graphToken) {
    try {
      const shareToken = encodeShareUrl(att.contentUrl);
      const graphUrl = `https://graph.microsoft.com/v1.0/shares/${shareToken}/driveItem/content`;
      log.info(`downloading attachment through Graph sharing API: ${att.name ?? att.contentUrl.slice(0, 80)}`);
      const resp = await fetch(graphUrl, {
        headers: { Authorization: `Bearer ${graphToken}` },
        redirect: "follow",
      });
      if (resp.ok) {
        const buf = Buffer.from(await resp.arrayBuffer());
        return {
          content: buf,
          contentType: resp.headers.get("content-type") ?? att.contentType,
        };
      }
      log.error(`Graph sharing download failed: ${resp.status} ${resp.statusText} for ${att.name ?? "(unnamed)"}`);
    } catch (err) {
      log.error(`Graph sharing download error: ${String(err)}`);
    }
  }

  if (att.contentUrl?.startsWith("https://") && graphToken) {
    try {
      const resp = await fetch(att.contentUrl, {
        headers: { Authorization: `Bearer ${graphToken}` },
      });
      if (resp.ok) {
        const buf = Buffer.from(await resp.arrayBuffer());
        return {
          content: buf,
          contentType: resp.headers.get("content-type") ?? att.contentType,
        };
      }
      log.error(`direct attachment fetch failed: ${resp.status} for ${att.contentUrl.slice(0, 120)}`);
    } catch (err) {
      log.error(`direct attachment fetch error: ${String(err)}`);
    }
  }

  if (att.content && (Buffer.isBuffer(att.content) || att.content instanceof ArrayBuffer)) {
    return {
      content: Buffer.from(att.content as ArrayBuffer),
      contentType: att.contentType ?? "application/octet-stream",
    };
  }

  return undefined;
}

export async function downloadInboundAttachments(
  context: unknown,
  log: AttachmentsLog,
  graphToken?: string,
): Promise<DownloadedAttachments> {
  const ctx = context as {
    activity: { attachments?: Attachment[]; conversation?: { id: string }; id?: string };
  };
  const attachments = ctx.activity.attachments ?? [];

  if (attachments.length === 0) {
    return { mediaPaths: [], mediaTypes: [] };
  }

  const files: Array<{ file: DownloadedFile; originalName?: string }> = [];
  for (const att of attachments) {
    const file = await downloadSingleAttachment(att, graphToken, log);
    if (file) {
      files.push({ file, originalName: att.name });
    }
  }

  if (files.length === 0) {
    log.info("no downloadable attachments found");
    return { mediaPaths: [], mediaTypes: [] };
  }

  const convId = sanitizePathSegment(ctx.activity.conversation?.id ?? "unknown");
  const activityId = sanitizePathSegment(ctx.activity.id ?? randomUUID());
  const mediaDir = join(CACHE_ROOT, convId, activityId);
  await mkdir(mediaDir, { recursive: true });

  const mediaPaths: string[] = [];
  const mediaTypes: string[] = [];

  for (let i = 0; i < files.length; i++) {
    const { file, originalName } = files[i];
    const filename = pickFilename(i, originalName, file.contentType);
    const target = join(mediaDir, filename);
    try {
      await writeFile(target, file.content);
      mediaPaths.push(target);
      mediaTypes.push(file.contentType);
    } catch (err) {
      log.error(`failed to write attachment ${filename}: ${String(err)}`);
    }
  }

  log.info(`downloaded ${mediaPaths.length} attachment(s) to ${mediaDir}`);
  return { mediaPaths, mediaTypes, mediaDir };
}

export async function reapStaleAttachmentCache(log: AttachmentsLog): Promise<void> {
  let convDirs: string[];
  try {
    convDirs = await readdir(CACHE_ROOT);
  } catch {
    return;
  }

  const now = Date.now();
  for (const conv of convDirs) {
    const convPath = join(CACHE_ROOT, conv);
    let activityDirs: string[];
    try {
      activityDirs = await readdir(convPath);
    } catch {
      continue;
    }

    let survivors = 0;
    for (const activity of activityDirs) {
      const activityPath = join(convPath, activity);
      try {
        const st = await stat(activityPath);
        if (now - st.mtimeMs > TTL_MS) {
          await rm(activityPath, { recursive: true, force: true });
          log.debug?.("reaped stale attachment dir", { activityPath });
        } else {
          survivors++;
        }
      } catch (err) {
        log.debug?.("reaper skipped entry", { activityPath, error: String(err) });
      }
    }

    if (survivors === 0) {
      try {
        await rm(convPath, { recursive: true, force: true });
      } catch {
        // ignore cleanup races
      }
    }
  }
}
