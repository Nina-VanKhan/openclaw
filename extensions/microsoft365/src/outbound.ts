/**
 * Microsoft 365 Mail outbound handler.
 *
 * Sends email replies via Microsoft Graph API.
 */

import type { ChannelOutboundAdapter } from "openclaw/plugin-sdk";
import { GraphClient, resolveCredentials } from "./graph-client.js";
import type { Microsoft365Config } from "./types.js";

export const microsoft365Outbound: ChannelOutboundAdapter = {
  deliveryMode: "gateway",
  textChunkLimit: 10000,

  resolveTarget: ({ to, allowFrom, mode }) => {
    const trimmed = to?.trim().toLowerCase() ?? "";

    if (trimmed && trimmed.includes("@")) {
      return { ok: true, to: trimmed };
    }

    // Fall back to first allowFrom entry
    const allowList = (allowFrom ?? [])
      .map((entry) => String(entry).trim().toLowerCase())
      .filter(Boolean)
      .filter((entry) => entry !== "*");

    if (allowList.length > 0) {
      return { ok: true, to: allowList[0] };
    }

    return {
      ok: false,
      error: new Error(
        "Microsoft 365 Mail: no target email address. Provide `to=<email>` or set channels.microsoft365.allowFrom.",
      ),
    };
  },

  sendText: async ({ cfg, to, text, replyToId, threadId }) => {
    const m365 = cfg.channels?.microsoft365 as Microsoft365Config | undefined;
    const credentials = resolveCredentials(m365);
    if (!credentials?.refreshToken) {
      return { channel: "microsoft365", error: "Not configured (missing refreshToken)" };
    }

    const client = new GraphClient({ credentials });

    // If replying to a specific message, use the reply endpoint
    if (replyToId && typeof replyToId === "string") {
      await client.replyToMail(replyToId, text);
      return { channel: "microsoft365" };
    }

    // Otherwise send a new email
    await client.sendMail({
      to,
      subject: "Re: OpenClaw",
      body: text,
    });
    return { channel: "microsoft365" };
  },
};
