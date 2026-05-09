import { defineChannelPluginEntry } from "openclaw/plugin-sdk/channel-core";
import { a365Plugin } from "./src/channel.js";
import { setA365Runtime } from "./src/runtime.js";

// Re-export monitor for external use
export { monitorA365Provider } from "./src/monitor.js";
export { createGraphTools } from "./src/graph-tools.js";
export { sendMessageA365, sendAdaptiveCardA365 } from "./src/outbound.js";
export {
  saveConversationReference,
  getConversationReference,
  getConversationReferenceByUser,
  deleteConversationReference,
  listConversationReferences,
  clearConversationReferences,
} from "./src/conversation-store.js";
export type { A365Config, A365MessageMetadata, GraphCalendarEvent } from "./src/types.js";

const plugin: any = defineChannelPluginEntry({
  id: "a365",
  name: "Microsoft 365 Agents",
  description: "A365 channel plugin with native Graph API tools for calendar and email",
  plugin: a365Plugin,
  setRuntime: setA365Runtime,
});

export default plugin;
