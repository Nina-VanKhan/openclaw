/**
 * Microsoft Graph API Client
 *
 * Handles OAuth2 token management and Graph API calls for mail operations.
 */

import type {
  Microsoft365Config,
  Microsoft365Credentials,
  Microsoft365TokenResponse,
  GraphMailMessage,
  GraphSubscription,
  SendMailOptions,
} from "./types.js";
import { getSecret, SECRETS } from "./secrets.js";

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const AUTH_BASE = "https://login.microsoftonline.com";

// Required scopes for mail operations
export const MAIL_SCOPES = [
  "offline_access",
  "Mail.Read",
  "Mail.Send",
  "Mail.ReadWrite",
  "User.Read",
];

export type GraphClientConfig = {
  credentials: Microsoft365Credentials;
  onTokenRefresh?: (tokens: { accessToken: string; refreshToken?: string; expiresAt: number }) => void;
};

export class GraphClient {
  private credentials: Microsoft365Credentials;
  private accessToken: string | undefined;
  private tokenExpiresAt: number = 0;
  private onTokenRefresh?: GraphClientConfig["onTokenRefresh"];

  constructor(config: GraphClientConfig) {
    this.credentials = config.credentials;
    this.accessToken = config.credentials.accessToken;
    this.tokenExpiresAt = config.credentials.tokenExpiresAt ?? 0;
    this.onTokenRefresh = config.onTokenRefresh;
  }

  /**
   * Get a valid access token, refreshing if necessary
   */
  async getAccessToken(): Promise<string> {
    const now = Date.now();
    const buffer = 5 * 60 * 1000; // 5 minute buffer

    if (this.accessToken && this.tokenExpiresAt > now + buffer) {
      return this.accessToken;
    }

    if (!this.credentials.refreshToken) {
      throw new Error("No refresh token available. Run auth flow first.");
    }

    const tokens = await this.refreshAccessToken(this.credentials.refreshToken);
    this.accessToken = tokens.access_token;
    this.tokenExpiresAt = now + tokens.expires_in * 1000;

    if (tokens.refresh_token) {
      this.credentials.refreshToken = tokens.refresh_token;
    }

    this.onTokenRefresh?.({
      accessToken: this.accessToken,
      refreshToken: this.credentials.refreshToken,
      expiresAt: this.tokenExpiresAt,
    });

    return this.accessToken;
  }

  /**
   * Refresh the access token using the refresh token
   */
  private async refreshAccessToken(refreshToken: string): Promise<Microsoft365TokenResponse> {
    const tenant = this.credentials.tenantId || "common";
    const url = `${AUTH_BASE}/${tenant}/oauth2/v2.0/token`;

    const params: Record<string, string> = {
      client_id: this.credentials.clientId,
      refresh_token: refreshToken,
      grant_type: "refresh_token",
      scope: MAIL_SCOPES.join(" "),
    };

    // Only include client_secret for confidential clients (not public/native apps)
    if (this.credentials.clientSecret) {
      params.client_secret = this.credentials.clientSecret;
    }

    const body = new URLSearchParams(params);

    const res = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: body.toString(),
    });

    if (!res.ok) {
      const error = await res.text();
      throw new Error(`Token refresh failed: ${res.status} ${error}`);
    }

    return res.json() as Promise<Microsoft365TokenResponse>;
  }

  /**
   * Make an authenticated request to Microsoft Graph
   */
  async request<T>(
    method: "GET" | "POST" | "PATCH" | "DELETE",
    endpoint: string,
    body?: unknown,
  ): Promise<T> {
    const token = await this.getAccessToken();
    const url = endpoint.startsWith("http") ? endpoint : `${GRAPH_BASE}${endpoint}`;

    const res = await fetch(url, {
      method,
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: body ? JSON.stringify(body) : undefined,
    });

    if (!res.ok) {
      const error = await res.text();
      throw new Error(`Graph API error: ${res.status} ${error}`);
    }

    if (res.status === 204) {
      return undefined as T;
    }

    return res.json() as Promise<T>;
  }

  /**
   * Get current user profile
   */
  async getMe(): Promise<{ id: string; mail: string; displayName: string; userPrincipalName: string }> {
    return this.request("GET", "/me");
  }

  /**
   * List messages in a folder
   */
  async listMessages(options?: {
    folderId?: string;
    top?: number;
    filter?: string;
    orderBy?: string;
    select?: string[];
  }): Promise<{ value: GraphMailMessage[]; "@odata.nextLink"?: string }> {
    const folder = options?.folderId ?? "inbox";
    const params = new URLSearchParams();

    if (options?.top) params.set("$top", String(options.top));
    if (options?.filter) params.set("$filter", options.filter);
    if (options?.orderBy) params.set("$orderby", options.orderBy);
    if (options?.select) params.set("$select", options.select.join(","));

    const query = params.toString();
    const endpoint = `/me/mailFolders/${folder}/messages${query ? `?${query}` : ""}`;

    return this.request("GET", endpoint);
  }

  /**
   * Get a specific message
   */
  async getMessage(messageId: string): Promise<GraphMailMessage> {
    return this.request("GET", `/me/messages/${messageId}`);
  }

  /**
   * Mark a message as read
   */
  async markAsRead(messageId: string): Promise<void> {
    await this.request("PATCH", `/me/messages/${messageId}`, { isRead: true });
  }

  /**
   * Send an email
   */
  async sendMail(options: SendMailOptions): Promise<void> {
    const toRecipients = (Array.isArray(options.to) ? options.to : [options.to]).map((addr) => ({
      emailAddress: { address: addr },
    }));

    const ccRecipients = options.cc
      ? (Array.isArray(options.cc) ? options.cc : [options.cc]).map((addr) => ({
          emailAddress: { address: addr },
        }))
      : undefined;

    const bccRecipients = options.bcc
      ? (Array.isArray(options.bcc) ? options.bcc : [options.bcc]).map((addr) => ({
          emailAddress: { address: addr },
        }))
      : undefined;

    const message: Record<string, unknown> = {
      subject: options.subject,
      body: {
        contentType: options.bodyType === "html" ? "HTML" : "Text",
        content: options.body,
      },
      toRecipients,
    };

    if (ccRecipients) message.ccRecipients = ccRecipients;
    if (bccRecipients) message.bccRecipients = bccRecipients;
    if (options.importance) message.importance = options.importance;
    if (options.replyTo) {
      message.replyTo = [{ emailAddress: { address: options.replyTo } }];
    }
    if (options.attachments) {
      message.attachments = options.attachments.map((att) => ({
        "@odata.type": "#microsoft.graph.fileAttachment",
        name: att.name,
        contentType: att.contentType,
        contentBytes: att.contentBytes,
      }));
    }

    await this.request("POST", "/me/sendMail", { message, saveToSentItems: true });
  }

  /**
   * Reply to an email
   */
  async replyToMail(messageId: string, body: string, replyAll?: boolean): Promise<void> {
    const endpoint = `/me/messages/${messageId}/${replyAll ? "replyAll" : "reply"}`;
    await this.request("POST", endpoint, {
      message: {
        body: {
          contentType: "Text",
          content: body,
        },
      },
    });
  }

  /**
   * Create a webhook subscription for new mail
   */
  async createSubscription(options: {
    notificationUrl: string;
    clientState?: string;
    expirationMinutes?: number;
  }): Promise<GraphSubscription> {
    const expirationDateTime = new Date(
      Date.now() + (options.expirationMinutes ?? 4230) * 60 * 1000, // Max ~3 days
    ).toISOString();

    return this.request("POST", "/subscriptions", {
      changeType: "created",
      notificationUrl: options.notificationUrl,
      resource: "/me/mailFolders/inbox/messages",
      expirationDateTime,
      clientState: options.clientState,
    });
  }

  /**
   * Renew a webhook subscription
   */
  async renewSubscription(subscriptionId: string, expirationMinutes?: number): Promise<GraphSubscription> {
    const expirationDateTime = new Date(
      Date.now() + (expirationMinutes ?? 4230) * 60 * 1000,
    ).toISOString();

    return this.request("PATCH", `/subscriptions/${subscriptionId}`, {
      expirationDateTime,
    });
  }

  /**
   * Delete a webhook subscription
   */
  async deleteSubscription(subscriptionId: string): Promise<void> {
    await this.request("DELETE", `/subscriptions/${subscriptionId}`);
  }

  /**
   * List active subscriptions
   */
  async listSubscriptions(): Promise<{ value: GraphSubscription[] }> {
    return this.request("GET", "/subscriptions");
  }
}

/**
 * Extract credentials from config
 */
/**
 * Resolve credentials from config and secrets store.
 *
 * Client secret is resolved in order:
 * 1. From config.clientSecret (if set directly)
 * 2. From secrets store (MICROSOFT365_CLIENT_SECRET)
 *
 * This allows users to either set the secret in config or use the secure secrets store.
 */
export function resolveCredentials(config?: Microsoft365Config): Microsoft365Credentials | null {
  if (!config?.clientId) {
    return null;
  }

  // Resolve client secret: config first, then secrets store
  const clientSecret = config.clientSecret || getSecret(SECRETS.CLIENT_SECRET);
  if (!clientSecret) {
    return null;
  }

  return {
    clientId: config.clientId,
    clientSecret,
    tenantId: config.tenantId ?? "common",
    refreshToken: config.refreshToken,
    accessToken: config.accessToken,
    tokenExpiresAt: config.tokenExpiresAt,
  };
}

/**
 * Build OAuth2 authorization URL for user consent
 */
export function buildAuthUrl(config: {
  clientId: string;
  tenantId?: string;
  redirectUri: string;
  state?: string;
}): string {
  const tenant = config.tenantId ?? "common";
  const params = new URLSearchParams({
    client_id: config.clientId,
    response_type: "code",
    redirect_uri: config.redirectUri,
    response_mode: "query",
    scope: MAIL_SCOPES.join(" "),
    state: config.state ?? "openclaw-microsoft365",
  });

  return `${AUTH_BASE}/${tenant}/oauth2/v2.0/authorize?${params.toString()}`;
}

/**
 * Exchange authorization code for tokens
 */
export async function exchangeCodeForTokens(config: {
  clientId: string;
  clientSecret: string;
  tenantId?: string;
  code: string;
  redirectUri: string;
}): Promise<Microsoft365TokenResponse> {
  const tenant = config.tenantId ?? "common";
  const url = `${AUTH_BASE}/${tenant}/oauth2/v2.0/token`;

  const body = new URLSearchParams({
    client_id: config.clientId,
    client_secret: config.clientSecret,
    code: config.code,
    redirect_uri: config.redirectUri,
    grant_type: "authorization_code",
    scope: MAIL_SCOPES.join(" "),
  });

  const res = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: body.toString(),
  });

  if (!res.ok) {
    const error = await res.text();
    throw new Error(`Token exchange failed: ${res.status} ${error}`);
  }

  return res.json() as Promise<Microsoft365TokenResponse>;
}
