/**
 * Microsoft 365 Mail Channel Plugin
 *
 * Provides email as a channel for OpenClaw.
 */

import { z } from "zod";
import type { ChannelPlugin, OpenClawConfig, RuntimeEnv } from "openclaw/plugin-sdk";
import { buildChannelConfigSchema, DEFAULT_ACCOUNT_ID, PAIRING_APPROVED_MESSAGE, logInboundDrop } from "openclaw/plugin-sdk";

import type { Microsoft365Config, Microsoft365AccountSnapshot, GraphMailMessage } from "./types.js";
import { getMicrosoft365Runtime } from "./runtime.js";

/**
 * Zod schema for Microsoft 365 channel configuration
 */
const Microsoft365ConfigSchema = z.object({
  enabled: z.boolean().optional(),
  clientId: z.string().optional(),
  clientSecret: z.string().optional(),
  tenantId: z.string().optional(),
  refreshToken: z.string().optional(),
  accessToken: z.string().optional(),
  tokenExpiresAt: z.number().optional(),
  userEmail: z.string().optional(),
  webhook: z.object({
    port: z.number().optional(),
    path: z.string().optional(),
    publicUrl: z.string().optional(),
  }).optional(),
  pollIntervalMs: z.number().optional(),
  folders: z.array(z.string()).optional(),
  allowFrom: z.array(z.string()).optional(),
  dmPolicy: z.enum(["open", "pairing", "allowlist"]).optional(),
});
import { GraphClient, resolveCredentials } from "./graph-client.js";
import { startMailMonitor, type MailMonitorRuntime } from "./monitor.js";
import { microsoft365Outbound } from "./outbound.js";

type ResolvedMicrosoft365Account = {
  accountId: string;
  enabled: boolean;
  configured: boolean;
  userEmail?: string;
};

// Runtime state (module-level for now)
let runtimeState: {
  connected: boolean;
  webhookActive: boolean;
  subscriptionId: string | null;
  lastError: string | null;
  userEmail: string | null;
} = {
  connected: false,
  webhookActive: false,
  subscriptionId: null,
  lastError: null,
  userEmail: null,
};

const meta = {
  id: "microsoft365",
  label: "Microsoft 365 Mail",
  selectionLabel: "Microsoft 365 Mail (Outlook)",
  docsPath: "/channels/microsoft365",
  docsLabel: "microsoft365",
  blurb: "Email via Microsoft Graph API with webhooks.",
  aliases: ["outlook", "office365", "o365", "email"],
  order: 61,
} as const;

export const microsoft365Plugin: ChannelPlugin<ResolvedMicrosoft365Account> = {
  id: "microsoft365",
  meta: { ...meta },

  capabilities: {
    chatTypes: ["direct"], // Email is essentially DM
    polls: false,
    reactions: false,
    edit: false,
    unsend: false,
    reply: true, // Email supports replies
    effects: false,
    groupManagement: false,
    threads: true, // Email threads via conversationId
    media: true, // Attachments
  },

  agentPrompt: {
    messageToolHints: () => [
      "- Email targeting: use email addresses directly (e.g., `to=user@example.com`).",
      "- Reply to current thread: omit `to` to reply to the incoming email.",
      "- Email formatting: HTML is supported via `bodyType=html`.",
      "- Attachments: provide base64-encoded content in `attachments` array.",
    ],
  },

  reload: { configPrefixes: ["channels.microsoft365"] },

  configSchema: buildChannelConfigSchema(Microsoft365ConfigSchema),

  config: {
    listAccountIds: () => [DEFAULT_ACCOUNT_ID],

    resolveAccount: (cfg) => {
      const m365 = cfg.channels?.microsoft365 as Microsoft365Config | undefined;
      const credentials = resolveCredentials(m365);
      return {
        accountId: DEFAULT_ACCOUNT_ID,
        enabled: m365?.enabled !== false,
        configured: Boolean(credentials?.refreshToken),
        userEmail: m365?.userEmail ?? runtimeState.userEmail ?? undefined,
      };
    },

    defaultAccountId: () => DEFAULT_ACCOUNT_ID,

    setAccountEnabled: ({ cfg, enabled }) => ({
      ...cfg,
      channels: {
        ...cfg.channels,
        microsoft365: {
          ...cfg.channels?.microsoft365,
          enabled,
        },
      },
    }),

    deleteAccount: ({ cfg }) => {
      const next = { ...cfg } as OpenClawConfig;
      const nextChannels = { ...cfg.channels };
      delete nextChannels.microsoft365;
      if (Object.keys(nextChannels).length > 0) {
        next.channels = nextChannels;
      } else {
        delete next.channels;
      }
      return next;
    },

    isConfigured: (_account, cfg) => {
      const m365 = cfg.channels?.microsoft365 as Microsoft365Config | undefined;
      return Boolean(resolveCredentials(m365)?.refreshToken);
    },

    describeAccount: (account) => ({
      accountId: account.accountId,
      enabled: account.enabled,
      configured: account.configured,
      userEmail: account.userEmail,
    }),

    resolveAllowFrom: ({ cfg }) => {
      const m365 = cfg.channels?.microsoft365 as Microsoft365Config | undefined;
      return m365?.allowFrom ?? [];
    },

    formatAllowFrom: ({ allowFrom }) =>
      allowFrom
        .map((entry) => String(entry).trim().toLowerCase())
        .filter(Boolean),
  },

  pairing: {
    idLabel: "email",
    normalizeAllowEntry: (entry) => entry.toLowerCase().trim(),
    notifyApproval: async ({ cfg, id }) => {
      const m365 = cfg.channels?.microsoft365 as Microsoft365Config | undefined;
      const credentials = resolveCredentials(m365);
      if (!credentials?.refreshToken) return;

      const client = new GraphClient({ credentials });
      await client.sendMail({
        to: id,
        subject: "OpenClaw Access Approved",
        body: PAIRING_APPROVED_MESSAGE,
      });
    },
  },

  messaging: {
    normalizeTarget: (raw) => {
      const trimmed = raw.trim().toLowerCase();
      // Remove common prefixes
      if (trimmed.startsWith("email:")) return trimmed.slice(6).trim();
      if (trimmed.startsWith("mailto:")) return trimmed.slice(7).trim();
      // Validate email format (basic check)
      if (trimmed.includes("@")) return trimmed;
      return undefined;
    },
    targetResolver: {
      looksLikeId: (raw) => raw.includes("@"),
      hint: "<email@address.com>",
    },
  },

  directory: {
    self: async () => null,
    listPeers: async ({ cfg, query, limit }) => {
      const m365 = cfg.channels?.microsoft365 as Microsoft365Config | undefined;
      const allowFrom = m365?.allowFrom ?? [];
      const q = query?.trim().toLowerCase() || "";

      return allowFrom
        .map((email) => String(email).trim().toLowerCase())
        .filter(Boolean)
        .filter((email) => (q ? email.includes(q) : true))
        .slice(0, limit && limit > 0 ? limit : undefined)
        .map((id) => ({ kind: "user", id }) as const);
    },
    listGroups: async () => [], // Email doesn't have groups in the same sense
  },

  security: {
    collectWarnings: ({ cfg }) => {
      const m365 = cfg.channels?.microsoft365 as Microsoft365Config | undefined;
      const dmPolicy = m365?.dmPolicy ?? "pairing";

      if (dmPolicy === "open") {
        return [
          "- Microsoft 365 Mail: dmPolicy=\"open\" allows anyone to send emails that trigger the agent. Consider using \"pairing\" or \"allowlist\".",
        ];
      }
      return [];
    },
  },

  status: {
    defaultRuntime: {
      accountId: DEFAULT_ACCOUNT_ID,
      running: false,
      lastStartAt: null,
      lastStopAt: null,
      lastError: null,
    },

    buildChannelSummary: ({ snapshot }) => ({
      configured: snapshot.configured ?? false,
      running: snapshot.connected ?? false,
      lastStartAt: snapshot.lastStartAt ?? null,
      lastStopAt: snapshot.lastStopAt ?? null,
      lastError: snapshot.lastError ?? null,
      userEmail: snapshot.userEmail ?? null,
      webhookActive: snapshot.webhookActive ?? false,
      subscriptionId: snapshot.subscriptionId ?? null,
    }),

    probeAccount: async ({ cfg }) => {
      const m365 = cfg.channels?.microsoft365 as Microsoft365Config | undefined;
      const credentials = resolveCredentials(m365);
      if (!credentials?.refreshToken) {
        return { ok: false, error: "Not configured" };
      }

      try {
        const client = new GraphClient({ credentials });
        const me = await client.getMe();
        return { ok: true, email: me.mail || me.userPrincipalName };
      } catch (err) {
        return { ok: false, error: err instanceof Error ? err.message : String(err) };
      }
    },

    buildAccountSnapshot: ({ account, runtime, probe }) => ({
      accountId: account.accountId,
      enabled: account.enabled,
      configured: account.configured,
      userEmail: account.userEmail ?? runtimeState.userEmail ?? undefined,
      connected: runtimeState.connected,
      webhookActive: runtimeState.webhookActive,
      subscriptionId: runtimeState.subscriptionId,
      running: runtime?.running ?? false,
      lastStartAt: runtime?.lastStartAt ?? null,
      lastStopAt: runtime?.lastStopAt ?? null,
      lastError: runtimeState.lastError ?? runtime?.lastError ?? null,
      probe,
    }),
  },

  gateway: {
    startAccount: async (ctx) => {
      const { cfg, runtime, abortSignal, accountId } = ctx;
      const m365 = cfg.channels?.microsoft365 as Microsoft365Config | undefined;

      ctx.log?.info("Starting Microsoft 365 Mail monitor");
      ctx.setStatus({ accountId, running: true, lastStartAt: Date.now() });

      const monitorRuntime: MailMonitorRuntime = {
        info: (msg) => ctx.log?.info(msg),
        warn: (msg) => ctx.log?.warn(msg),
        error: (msg) => ctx.log?.error(msg),
        debug: (msg) => ctx.log?.debug?.(msg),
      };

      try {
        const stopMonitor = await startMailMonitor({
          cfg,
          runtime: monitorRuntime,
          abortSignal,
          onMail: async ({ message }) => {
            await handleIncomingMail({
              message,
              cfg,
              runtime,
              accountId,
              log: ctx.log,
            });
          },
          onStatusChange: (status) => {
            runtimeState.connected = status.connected;
            runtimeState.webhookActive = status.webhookActive;
            runtimeState.subscriptionId = status.subscriptionId ?? null;
            runtimeState.lastError = status.error ?? null;
          },
        });

        // Return cleanup function
        return async () => {
          ctx.log?.info("Stopping Microsoft 365 Mail monitor");
          await stopMonitor();
          ctx.setStatus({ accountId, running: false, lastStopAt: Date.now() });
          runtimeState.connected = false;
        };
      } catch (err) {
        const msg = err instanceof Error ? err.message : String(err);
        ctx.log?.error(`Failed to start mail monitor: ${msg}`);
        runtimeState.lastError = msg;
        ctx.setStatus({ accountId, running: false, lastError: msg });
        throw err;
      }
    },
  },

  outbound: microsoft365Outbound,
};

const CHANNEL_ID = "microsoft365" as const;

/**
 * Handle an incoming email message using the proper OpenClaw dispatch pipeline.
 */
async function handleIncomingMail(params: {
  message: GraphMailMessage;
  cfg: OpenClawConfig;
  runtime: RuntimeEnv;
  accountId: string;
  log?: { info: (msg: string) => void; warn: (msg: string) => void; error: (msg: string) => void; debug?: (msg: string) => void };
}): Promise<void> {
  const { message, cfg, runtime, accountId, log } = params;
  const core = getMicrosoft365Runtime();

  const fromEmail = message.from?.emailAddress?.address?.toLowerCase();
  const fromName = message.from?.emailAddress?.name;
  const subject = message.subject ?? "(no subject)";
  const rawBody = message.body?.contentType === "text"
    ? message.body.content?.trim()
    : message.bodyPreview?.trim() ?? "";

  if (!fromEmail) {
    log?.warn?.("microsoft365: dropping message with no sender address");
    return;
  }

  log?.info?.(`Incoming email from ${fromEmail}: ${subject}`);

  // --- DM policy enforcement ---
  const m365 = cfg.channels?.microsoft365 as Microsoft365Config | undefined;
  const dmPolicy = m365?.dmPolicy ?? "pairing";

  const configAllowFrom = (m365?.allowFrom ?? []).map((e) => String(e).trim().toLowerCase()).filter(Boolean);
  const storeAllowFrom = await core.channel.pairing.readAllowFromStore(CHANNEL_ID).catch(() => []);
  const storeAllowList = storeAllowFrom.map((e) => String(e).trim().toLowerCase()).filter(Boolean);
  const effectiveAllowFrom = [...configAllowFrom, ...storeAllowList];

  if (dmPolicy !== "open") {
    const senderAllowed = effectiveAllowFrom.some(
      (entry) => entry === fromEmail || entry === "*",
    );
    if (!senderAllowed) {
      if (dmPolicy === "pairing") {
        const { code, created } = await core.channel.pairing.upsertPairingRequest({
          channel: CHANNEL_ID,
          id: fromEmail,
          meta: { name: fromName || undefined },
        });
        if (created) {
          // Send pairing reply via email
          try {
            const credentials = resolveCredentials(m365);
            if (credentials?.refreshToken) {
              const client = new GraphClient({ credentials });
              await client.replyToMail(
                message.id,
                core.channel.pairing.buildPairingReply({
                  channel: CHANNEL_ID,
                  idLine: `Your email address: ${fromEmail}`,
                  code,
                }),
              );
            }
          } catch (err) {
            log?.error?.(`microsoft365: pairing reply failed for ${fromEmail}: ${String(err)}`);
          }
        }
      }
      logInboundDrop({
        log: (msg) => log?.info?.(msg),
        channel: CHANNEL_ID,
        reason: `dmPolicy=${dmPolicy}`,
        target: fromEmail,
      });
      return;
    }
  }

  // Mark email as read so we don't process it again
  try {
    const credentials = resolveCredentials(m365);
    if (credentials?.refreshToken) {
      const client = new GraphClient({ credentials });
      await client.markAsRead(message.id);
    }
  } catch (err) {
    log?.warn?.(`microsoft365: failed to mark message as read: ${String(err)}`);
  }

  // --- Route resolution ---
  const route = core.channel.routing.resolveAgentRoute({
    cfg,
    channel: CHANNEL_ID,
    accountId,
    peer: { kind: "dm", id: fromEmail },
  });

  const storePath = core.channel.session.resolveStorePath(cfg.session?.store, {
    agentId: route.agentId,
  });

  // --- Build message body with envelope ---
  const envelopeOptions = core.channel.reply.resolveEnvelopeFormatOptions(cfg);
  const previousTimestamp = core.channel.session.readSessionUpdatedAt({
    storePath,
    sessionKey: route.sessionKey,
  });
  const messageTimestamp = message.receivedDateTime
    ? new Date(message.receivedDateTime).getTime()
    : Date.now();

  const bodyText = subject
    ? `Subject: ${subject}\n\n${rawBody}`
    : rawBody;

  const body = core.channel.reply.formatAgentEnvelope({
    channel: "Microsoft 365 Mail",
    from: fromName ? `${fromName} <${fromEmail}>` : fromEmail,
    timestamp: messageTimestamp,
    previousTimestamp,
    envelope: envelopeOptions,
    body: bodyText,
  });

  // --- Finalize inbound context ---
  const ctxPayload = core.channel.reply.finalizeInboundContext({
    Body: body,
    RawBody: bodyText,
    CommandBody: bodyText,
    From: `microsoft365:${fromEmail}`,
    To: `microsoft365:${accountId}`,
    SessionKey: route.sessionKey,
    AccountId: route.accountId,
    ChatType: "direct" as const,
    ConversationLabel: fromName ? `${fromName} <${fromEmail}>` : fromEmail,
    SenderName: fromName || undefined,
    SenderId: fromEmail,
    Provider: CHANNEL_ID,
    Surface: CHANNEL_ID,
    MessageSid: message.id,
    Timestamp: messageTimestamp,
    OriginatingChannel: CHANNEL_ID,
    OriginatingTo: `microsoft365:${accountId}`,
  });

  // --- Record session ---
  await core.channel.session.recordInboundSession({
    storePath,
    sessionKey: ctxPayload.SessionKey ?? route.sessionKey,
    ctx: ctxPayload,
    onRecordError: (err) => {
      log?.error?.(`microsoft365: failed updating session meta: ${String(err)}`);
    },
  });

  // --- Dispatch reply ---
  await core.channel.reply.dispatchReplyWithBufferedBlockDispatcher({
    ctx: ctxPayload,
    cfg,
    dispatcherOptions: {
      deliver: async (payload) => {
        const text = (payload as { text?: string }).text ?? "";
        if (!text.trim()) return;

        const credentials = resolveCredentials(m365);
        if (!credentials?.refreshToken) {
          log?.error?.("microsoft365: cannot reply, missing refreshToken");
          return;
        }

        const client = new GraphClient({ credentials });
        // Reply to the original message to keep the thread
        await client.replyToMail(message.id, text);
      },
      onError: (err, info) => {
        log?.error?.(`microsoft365 ${info.kind} reply failed: ${String(err)}`);
      },
    },
  });
}
