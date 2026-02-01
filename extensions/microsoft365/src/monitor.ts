/**
 * Microsoft 365 Mail Monitor
 *
 * Monitors inbox for new messages via webhooks (preferred) or polling (fallback).
 */

import type { Express } from "express";
import type { OpenClawConfig } from "openclaw/plugin-sdk";
import type { GraphMailMessage, GraphWebhookNotification, Microsoft365Config } from "./types.js";
import { GraphClient, resolveCredentials } from "./graph-client.js";

export type MailMonitorRuntime = {
  info: (msg: string) => void;
  warn: (msg: string) => void;
  error: (msg: string) => void;
  debug?: (msg: string) => void;
};

export type MailHandler = (params: {
  message: GraphMailMessage;
  cfg: OpenClawConfig;
  runtime: MailMonitorRuntime;
}) => Promise<void>;

export type MonitorContext = {
  cfg: OpenClawConfig;
  runtime: MailMonitorRuntime;
  abortSignal?: AbortSignal;
  onMail: MailHandler;
  onStatusChange?: (status: {
    connected: boolean;
    webhookActive: boolean;
    subscriptionId?: string | null;
    error?: string | null;
  }) => void;
};

type MonitorState = {
  client: GraphClient | null;
  subscriptionId: string | null;
  subscriptionExpires: Date | null;
  lastMessageTime: Date | null;
  pollInterval: ReturnType<typeof setInterval> | null;
  renewInterval: ReturnType<typeof setInterval> | null;
  webhookActive: boolean;
  stopping: boolean;
};

/**
 * Start monitoring for new mail
 */
export async function startMailMonitor(ctx: MonitorContext): Promise<() => Promise<void>> {
  const { cfg, runtime, abortSignal, onMail, onStatusChange } = ctx;
  const m365Config = cfg.channels?.microsoft365 as Microsoft365Config | undefined;

  const state: MonitorState = {
    client: null,
    subscriptionId: null,
    subscriptionExpires: null,
    lastMessageTime: null,
    pollInterval: null,
    renewInterval: null,
    webhookActive: false,
    stopping: false,
  };

  const credentials = resolveCredentials(m365Config);
  if (!credentials) {
    throw new Error("Microsoft 365 credentials not configured");
  }

  // Create Graph client with token refresh callback
  state.client = new GraphClient({
    credentials,
    onTokenRefresh: (tokens) => {
      runtime.debug?.(`Token refreshed, expires at ${new Date(tokens.expiresAt).toISOString()}`);
      // TODO: Persist updated tokens to config
    },
  });

  // Verify connection
  try {
    const me = await state.client.getMe();
    runtime.info(`Connected as ${me.mail || me.userPrincipalName}`);
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    runtime.error(`Failed to connect: ${msg}`);
    onStatusChange?.({ connected: false, webhookActive: false, error: msg });
    throw err;
  }

  // Try to set up webhooks if public URL is configured
  const webhookUrl = m365Config?.webhook?.publicUrl;
  if (webhookUrl) {
    try {
      await setupWebhook(state, ctx, webhookUrl);
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      runtime.warn(`Webhook setup failed, falling back to polling: ${msg}`);
    }
  }

  // Fall back to polling if webhooks not available
  if (!state.webhookActive) {
    startPolling(state, ctx);
  }

  onStatusChange?.({
    connected: true,
    webhookActive: state.webhookActive,
    subscriptionId: state.subscriptionId,
  });

  // Handle abort signal
  abortSignal?.addEventListener("abort", () => {
    stopMonitor(state, ctx);
  });

  // Return cleanup function
  return async () => {
    await stopMonitor(state, ctx);
  };
}

/**
 * Set up webhook subscription
 */
async function setupWebhook(
  state: MonitorState,
  ctx: MonitorContext,
  webhookUrl: string,
): Promise<void> {
  const { runtime, cfg } = ctx;
  const m365Config = cfg.channels?.microsoft365 as Microsoft365Config | undefined;

  if (!state.client) {
    throw new Error("Client not initialized");
  }

  // Generate a unique client state for validation
  const clientState = `openclaw-${Date.now()}-${Math.random().toString(36).slice(2)}`;

  try {
    const subscription = await state.client.createSubscription({
      notificationUrl: webhookUrl,
      clientState,
    });

    state.subscriptionId = subscription.id;
    state.subscriptionExpires = new Date(subscription.expirationDateTime);
    state.webhookActive = true;

    runtime.info(`Webhook subscription created: ${subscription.id}`);
    runtime.info(`Expires: ${subscription.expirationDateTime}`);

    // Set up subscription renewal (renew 30 minutes before expiry)
    const renewalMs = state.subscriptionExpires.getTime() - Date.now() - 30 * 60 * 1000;
    state.renewInterval = setInterval(
      async () => {
        if (state.stopping || !state.client || !state.subscriptionId) return;

        try {
          const renewed = await state.client.renewSubscription(state.subscriptionId);
          state.subscriptionExpires = new Date(renewed.expirationDateTime);
          runtime.info(`Subscription renewed, expires: ${renewed.expirationDateTime}`);
        } catch (err) {
          const msg = err instanceof Error ? err.message : String(err);
          runtime.error(`Failed to renew subscription: ${msg}`);
          // Fall back to polling
          state.webhookActive = false;
          startPolling(state, ctx);
        }
      },
      Math.max(renewalMs, 60000),
    ); // At least 1 minute
  } catch (err) {
    throw err;
  }
}

/**
 * Start polling for new messages
 */
function startPolling(state: MonitorState, ctx: MonitorContext): void {
  const { cfg, runtime } = ctx;
  const m365Config = cfg.channels?.microsoft365 as Microsoft365Config | undefined;
  const intervalMs = m365Config?.pollIntervalMs ?? 60000;

  runtime.info(`Starting polling (interval: ${intervalMs}ms)`);

  // Initial poll
  pollForNewMessages(state, ctx);

  // Set up interval
  state.pollInterval = setInterval(() => {
    if (!state.stopping) {
      pollForNewMessages(state, ctx);
    }
  }, intervalMs);
}

/**
 * Poll for new messages since last check
 */
async function pollForNewMessages(state: MonitorState, ctx: MonitorContext): Promise<void> {
  const { cfg, runtime, onMail } = ctx;

  if (!state.client) return;

  try {
    const filter = state.lastMessageTime
      ? `receivedDateTime gt ${state.lastMessageTime.toISOString()}`
      : `isRead eq false`;

    const result = await state.client.listMessages({
      top: 25,
      filter,
      orderBy: "receivedDateTime desc",
    });

    if (result.value.length > 0) {
      runtime.debug?.(`Found ${result.value.length} new message(s)`);

      // Process oldest first
      const messages = result.value.reverse();
      for (const message of messages) {
        await onMail({ message, cfg, runtime });

        // Update last message time
        const receivedAt = new Date(message.receivedDateTime);
        if (!state.lastMessageTime || receivedAt > state.lastMessageTime) {
          state.lastMessageTime = receivedAt;
        }
      }
    }
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    runtime.error(`Poll failed: ${msg}`);
  }
}

/**
 * Handle incoming webhook notification
 */
export async function handleWebhookNotification(params: {
  notification: GraphWebhookNotification;
  ctx: MonitorContext;
  state: MonitorState;
}): Promise<void> {
  const { notification, ctx, state } = params;
  const { cfg, runtime, onMail } = ctx;

  if (!state.client) return;

  for (const change of notification.value) {
    if (change.changeType !== "created") continue;

    const messageId = change.resourceData?.id;
    if (!messageId) continue;

    try {
      const message = await state.client.getMessage(messageId);
      await onMail({ message, cfg, runtime });
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      runtime.error(`Failed to fetch message ${messageId}: ${msg}`);
    }
  }
}

/**
 * Stop the monitor
 */
async function stopMonitor(state: MonitorState, ctx: MonitorContext): Promise<void> {
  const { runtime } = ctx;
  state.stopping = true;

  // Clear intervals
  if (state.pollInterval) {
    clearInterval(state.pollInterval);
    state.pollInterval = null;
  }
  if (state.renewInterval) {
    clearInterval(state.renewInterval);
    state.renewInterval = null;
  }

  // Delete subscription
  if (state.subscriptionId && state.client) {
    try {
      await state.client.deleteSubscription(state.subscriptionId);
      runtime.info("Webhook subscription deleted");
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      runtime.warn(`Failed to delete subscription: ${msg}`);
    }
  }

  state.webhookActive = false;
  state.subscriptionId = null;
  runtime.info("Mail monitor stopped");
}

/**
 * Create Express routes for webhook handling
 */
export function createWebhookRoutes(params: {
  path: string;
  ctx: MonitorContext;
  state: MonitorState;
}): (app: Express) => void {
  return (app) => {
    const { path, ctx, state } = params;
    const { runtime } = ctx;

    // Webhook validation endpoint (Microsoft sends a validation token on subscription creation)
    app.post(path, async (req, res) => {
      // Handle validation request
      const validationToken = req.query?.validationToken;
      if (validationToken) {
        runtime.debug?.("Webhook validation request received");
        res.set("Content-Type", "text/plain");
        res.status(200).send(validationToken);
        return;
      }

      // Handle notification
      try {
        const notification = req.body as GraphWebhookNotification;
        runtime.debug?.(`Webhook notification: ${notification.value?.length ?? 0} changes`);

        // Respond immediately (Microsoft expects quick response)
        res.status(202).send();

        // Process notifications asynchronously
        await handleWebhookNotification({ notification, ctx, state });
      } catch (err) {
        const msg = err instanceof Error ? err.message : String(err);
        runtime.error(`Webhook handler error: ${msg}`);
        res.status(500).send();
      }
    });
  };
}
