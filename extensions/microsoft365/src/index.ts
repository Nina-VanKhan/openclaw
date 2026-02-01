/**
 * Microsoft 365 Mail Channel - Source exports
 */

export { microsoft365Plugin } from "./channel.js";
export { GraphClient, resolveCredentials, buildAuthUrl, exchangeCodeForTokens, MAIL_SCOPES } from "./graph-client.js";
export { startMailMonitor, handleWebhookNotification, createWebhookRoutes } from "./monitor.js";
export { getSecret, hasSecret, SECRETS } from "./secrets.js";
export type {
  Microsoft365Config,
  Microsoft365Credentials,
  GraphMailMessage,
  GraphSubscription,
  GraphWebhookNotification,
  SendMailOptions,
  Microsoft365AccountSnapshot,
} from "./types.js";
