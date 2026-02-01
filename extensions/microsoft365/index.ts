/**
 * OpenClaw Microsoft 365 Mail Channel Plugin
 *
 * Provides email as a channel via Microsoft Graph API.
 * Features:
 * - Real-time email notifications via webhooks (or polling fallback)
 * - Send/reply to emails
 * - OAuth2 authentication with token refresh
 *
 * Setup:
 * 1. Register an app in Azure AD (Entra ID)
 * 2. Configure client ID, secret, and tenant ID
 * 3. Run the OAuth flow to get refresh token
 * 4. (Optional) Configure webhook URL for real-time notifications
 */

import type { OpenClawPluginApi } from "openclaw/plugin-sdk";
import { emptyPluginConfigSchema } from "openclaw/plugin-sdk";

import { microsoft365Plugin } from "./src/channel.js";

export { GraphClient, resolveCredentials, buildAuthUrl, exchangeCodeForTokens, MAIL_SCOPES } from "./src/graph-client.js";
export { startMailMonitor } from "./src/monitor.js";
export { getSecret, hasSecret, SECRETS } from "./src/secrets.js";
export type { Microsoft365Config, GraphMailMessage, SendMailOptions } from "./src/types.js";

const plugin = {
  id: "microsoft365",
  name: "Microsoft 365 Mail",
  description: "Microsoft 365 Mail channel plugin (Graph API)",
  configSchema: emptyPluginConfigSchema(),
  register(api: OpenClawPluginApi) {
    api.registerChannel({ plugin: microsoft365Plugin });

    // Register CLI commands for OAuth flow
    api.registerCommands?.([
      {
        name: "microsoft365",
        description: "Microsoft 365 Mail configuration",
        subcommands: [
          {
            name: "auth",
            description: "Run OAuth2 authentication flow",
            action: async (ctx) => {
              const { buildAuthUrl, exchangeCodeForTokens } = await import("./src/graph-client.js");

              const cfg = ctx.cfg?.channels?.microsoft365 as Record<string, unknown> | undefined;
              const clientId = cfg?.clientId as string | undefined;
              const clientSecret = cfg?.clientSecret as string | undefined;
              const tenantId = cfg?.tenantId as string | undefined;

              if (!clientId || !clientSecret) {
                ctx.log?.error("Missing clientId or clientSecret in config");
                ctx.log?.info("Set channels.microsoft365.clientId and channels.microsoft365.clientSecret first.");
                return;
              }

              // Use localhost callback for CLI auth
              const redirectUri = "http://localhost:8765/callback";
              const authUrl = buildAuthUrl({ clientId, tenantId, redirectUri });

              ctx.log?.info("Open this URL in your browser to authenticate:");
              ctx.log?.info(authUrl);
              ctx.log?.info("");
              ctx.log?.info("After authorizing, paste the 'code' parameter from the redirect URL:");

              // In a real CLI, we'd start a local server to capture the callback
              // For now, user manually pastes the code
              const code = await ctx.prompt?.("Authorization code: ");
              if (!code) {
                ctx.log?.error("No code provided");
                return;
              }

              try {
                const tokens = await exchangeCodeForTokens({
                  clientId,
                  clientSecret,
                  tenantId,
                  code: code.trim(),
                  redirectUri,
                });

                ctx.log?.info("Authentication successful!");
                ctx.log?.info(`Access token expires in ${tokens.expires_in} seconds`);

                if (tokens.refresh_token) {
                  ctx.log?.info("");
                  ctx.log?.info("Add this to your config (channels.microsoft365.refreshToken):");
                  ctx.log?.info(tokens.refresh_token);
                }
              } catch (err) {
                const msg = err instanceof Error ? err.message : String(err);
                ctx.log?.error(`Authentication failed: ${msg}`);
              }
            },
          },
          {
            name: "test",
            description: "Test the connection",
            action: async (ctx) => {
              const { GraphClient, resolveCredentials } = await import("./src/graph-client.js");

              const cfg = ctx.cfg?.channels?.microsoft365 as Record<string, unknown> | undefined;
              const credentials = resolveCredentials(cfg as never);

              if (!credentials?.refreshToken) {
                ctx.log?.error("Not configured. Run `openclaw microsoft365 auth` first.");
                return;
              }

              try {
                const client = new GraphClient({ credentials });
                const me = await client.getMe();
                ctx.log?.info(`Connected as: ${me.displayName} <${me.mail || me.userPrincipalName}>`);

                // Try to list recent messages
                const messages = await client.listMessages({ top: 3 });
                ctx.log?.info(`Recent messages: ${messages.value.length}`);
                for (const msg of messages.value) {
                  ctx.log?.info(`  - ${msg.subject} (from ${msg.from?.emailAddress?.address})`);
                }
              } catch (err) {
                const msg = err instanceof Error ? err.message : String(err);
                ctx.log?.error(`Connection test failed: ${msg}`);
              }
            },
          },
        ],
      },
    ]);
  },
};

export default plugin;
