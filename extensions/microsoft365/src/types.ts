/**
 * Microsoft 365 Channel Types
 */

export type Microsoft365Config = {
  enabled?: boolean;
  clientId?: string;
  clientSecret?: string;
  tenantId?: string;
  refreshToken?: string;
  accessToken?: string;
  tokenExpiresAt?: number;
  userEmail?: string;
  webhook?: {
    port?: number;
    path?: string;
    publicUrl?: string;
  };
  pollIntervalMs?: number;
  folders?: string[];
  allowFrom?: string[];
  dmPolicy?: "open" | "pairing" | "allowlist";
};

export type Microsoft365Credentials = {
  clientId: string;
  clientSecret: string;
  tenantId: string;
  refreshToken?: string;
  accessToken?: string;
  tokenExpiresAt?: number;
};

export type Microsoft365TokenResponse = {
  access_token: string;
  refresh_token?: string;
  expires_in: number;
  token_type: string;
  scope: string;
};

export type GraphMailMessage = {
  id: string;
  createdDateTime: string;
  receivedDateTime: string;
  subject: string;
  bodyPreview: string;
  body: {
    contentType: "text" | "html";
    content: string;
  };
  from: {
    emailAddress: {
      name: string;
      address: string;
    };
  };
  toRecipients: Array<{
    emailAddress: {
      name: string;
      address: string;
    };
  }>;
  ccRecipients?: Array<{
    emailAddress: {
      name: string;
      address: string;
    };
  }>;
  replyTo?: Array<{
    emailAddress: {
      name: string;
      address: string;
    };
  }>;
  conversationId: string;
  conversationIndex: string;
  isRead: boolean;
  hasAttachments: boolean;
  internetMessageId: string;
  parentFolderId: string;
  importance: "low" | "normal" | "high";
};

export type GraphSubscription = {
  id: string;
  resource: string;
  changeType: string;
  notificationUrl: string;
  expirationDateTime: string;
  clientState?: string;
};

export type GraphWebhookNotification = {
  value: Array<{
    subscriptionId: string;
    clientState?: string;
    changeType: "created" | "updated" | "deleted";
    resource: string;
    resourceData?: {
      "@odata.type": string;
      "@odata.id": string;
      "@odata.etag"?: string;
      id: string;
    };
    subscriptionExpirationDateTime: string;
    tenantId: string;
  }>;
};

export type SendMailOptions = {
  to: string | string[];
  cc?: string | string[];
  bcc?: string | string[];
  subject: string;
  body: string;
  bodyType?: "text" | "html";
  replyTo?: string;
  importance?: "low" | "normal" | "high";
  attachments?: Array<{
    name: string;
    contentType: string;
    contentBytes: string; // base64
  }>;
};

export type Microsoft365AccountSnapshot = {
  accountId: string;
  enabled: boolean;
  configured: boolean;
  connected: boolean;
  userEmail?: string;
  lastMessageAt?: number | null;
  lastError?: string | null;
  webhookActive?: boolean;
  subscriptionId?: string | null;
  subscriptionExpires?: string | null;
};
