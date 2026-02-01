/**
 * Microsoft 365 Mail Channel Plugin
 *
 * Provides email as a channel for OpenClaw.
 */

import { z } from "zod";
import type { ChannelPlugin, OpenClawConfig } from "openclaw/plugin-sdk";
import { buildChannelConfigSchema, DEFAULT_ACCOUNT_ID, PAIRING_APPROVED_MESSAGE } from "openclaw/plugin-sdk";

import type { Microsoft365Config, Microsoft365AccountSnapshot, GraphMailMessage } from "./types.js";

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
          onMail: async ({ message, cfg: _cfg, runtime: _runtime }) => {
            await handleIncomingMail({ message, ctx });
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

  // outbound: microsoft365Outbound, // TODO: implement
};

/**
 * Handle an incoming email message
 */
async function handleIncomingMail(params: {
  message: GraphMailMessage;
  ctx: {
    cfg: OpenClawConfig;
    runtime: unknown;
    log?: { info: (msg: string) => void };
    injectMessage?: (msg: {
      channel: string;
      accountId: string;
      from: string;
      fromName?: string;
      to?: string;
      text: string;
      messageId?: string;
      threadId?: string;
      metadata?: Record<string, unknown>;
    }) => Promise<void>;
  };
}): Promise<void> {
  const { message, ctx } = params;

  const fromEmail = message.from?.emailAddress?.address;
  const fromName = message.from?.emailAddress?.name;
  const subject = message.subject;
  const body = message.body?.contentType === "text" ? message.body.content : message.bodyPreview;

  ctx.log?.info(`Incoming email from ${fromEmail}: ${subject}`);

  // Format message for injection
  const text = `Subject: ${subject}\n\n${body}`;

  // Inject into OpenClaw as a session event
  await ctx.injectMessage?.({
    channel: "microsoft365",
    accountId: DEFAULT_ACCOUNT_ID,
    from: fromEmail,
    fromName,
    text,
    messageId: message.id,
    threadId: message.conversationId,
    metadata: {
      subject,
      internetMessageId: message.internetMessageId,
      importance: message.importance,
      hasAttachments: message.hasAttachments,
    },
  });
}
