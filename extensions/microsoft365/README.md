# Microsoft 365 Mail Channel Plugin

Email as a channel for OpenClaw using Microsoft Graph API.

## Overview

This plugin enables OpenClaw to:
- Receive emails as incoming messages
- Send emails as replies
- Monitor inbox for new messages via polling or webhooks

## Prerequisites

- Microsoft 365 account (personal or organizational)
- Azure AD access to register an application
- Admin consent (for organizational accounts)

## Setup Guide

### Step 1: Register Azure AD Application

1. Go to [Azure Portal](https://portal.azure.com) → **Azure Active Directory** → **App registrations**
2. Click **New registration**
3. Configure:
   - **Name:** `OpenClaw Mail` (or your preferred name)
   - **Supported account types:** 
     - "Accounts in this organizational directory only" (single tenant) for org accounts
     - "Accounts in any organizational directory and personal Microsoft accounts" for broader access
   - **Redirect URI:** Select "Public client/native" and enter `http://localhost:8765/callback`
4. Click **Register**
5. Note down:
   - **Application (client) ID:** `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`
   - **Directory (tenant) ID:** `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`

### Step 2: Configure API Permissions

1. In your app registration, go to **API permissions**
2. Click **Add a permission** → **Microsoft Graph** → **Delegated permissions**
3. Add these permissions (search and select each):
   - `Mail.Read` — Read user mail
   - `Mail.ReadWrite` — Read and write user mail
   - `Mail.Send` — Send mail as the user
   - `User.Read` — Sign in and read user profile
   - `offline_access` — Maintain access (for refresh tokens)
4. Click **Add permissions**

**Note:** To select permissions in the Azure UI, click the row then press **Space** to toggle the checkbox.

### Step 3: Create Client Secret

1. Go to **Certificates & secrets** → **Client secrets**
2. Click **New client secret**
3. Add a description (e.g., "OpenClaw") and select expiry
4. Click **Add**
5. **Copy the secret value immediately** — it won't be shown again

### Step 4: Grant Admin Consent (Organizational Accounts)

For organizational Microsoft 365 accounts, an admin must consent:

**Option A: Via Azure Portal**
1. Go to **API permissions**
2. Click **Grant admin consent for [Your Organization]**
3. Confirm

**Option B: Via URL**
Navigate to (replace values):
```
https://login.microsoftonline.com/{tenant-id}/adminconsent?client_id={client-id}
```

**Option C: Via Enterprise Applications**
1. Azure Portal → **Enterprise applications**
2. Find your app → **Permissions**
3. Click **Grant admin consent**

### Step 5: Store Client Secret

**Option A: Use OpenClaw secrets store (recommended)**

```bash
openclaw secrets set MICROSOFT365_CLIENT_SECRET
# Paste your client secret when prompted
```

Optionally add a description:
```bash
openclaw secrets set MICROSOFT365_CLIENT_SECRET --description "Azure AD app for email"
```

The plugin automatically reads from the secrets store at `~/.openclaw/secrets.json`.

**Option B: Set in config directly (less secure)**

Add to `channels.microsoft365.clientSecret` in your config. Not recommended as the value is stored in plaintext.

### Step 6: Configure OpenClaw

Add to your `openclaw.json`:

```json
{
  "channels": {
    "microsoft365": {
      "enabled": true,
      "clientId": "your-client-id",
      "tenantId": "your-tenant-id",
      "dmPolicy": "allowlist",
      "allowFrom": [
        "your-email@example.com"
      ]
    }
  },
  "plugins": {
    "entries": {
      "microsoft365": {
        "enabled": true
      }
    }
  }
}
```

**Note:** The client secret is stored in secrets, not in config. The plugin reads it from `${secret:MICROSOFT365_CLIENT_SECRET}`.

### Step 7: Complete OAuth Flow

Run the authentication command:

```bash
openclaw microsoft365 auth
```

This will:
1. Open a browser to Microsoft login
2. Prompt you to sign in and consent
3. Redirect to localhost with an auth code
4. Exchange the code for tokens
5. Save the refresh token to config

### Step 8: Test Connection

```bash
openclaw microsoft365 test
```

This verifies:
- Token refresh works
- Can access Microsoft Graph API
- Can read mailbox

## Configuration Reference

| Field | Type | Description |
|-------|------|-------------|
| `enabled` | boolean | Enable/disable the channel |
| `clientId` | string | Azure AD Application (client) ID |
| `tenantId` | string | Azure AD Directory (tenant) ID |
| `dmPolicy` | string | `"allowlist"`, `"open"`, or `"pairing"` |
| `allowFrom` | string[] | Email addresses allowed to message (for allowlist mode) |
| `pollIntervalMs` | number | Polling interval in ms (default: 30000) |
| `userEmail` | string | (Auto-set) Authenticated user's email |
| `refreshToken` | string | (Auto-set) OAuth refresh token |

## Architecture

```
┌─────────────────┐      ┌──────────────────┐      ┌─────────────────┐
│  Microsoft 365  │ ←──→ │  Graph API       │ ←──→ │  OpenClaw       │
│  Mailbox        │      │  (REST + OAuth)  │      │  Gateway        │
└─────────────────┘      └──────────────────┘      └─────────────────┘
                                                           │
                                                           ▼
                                                   ┌─────────────────┐
                                                   │  Agent Session  │
                                                   └─────────────────┘
```

**Inbound flow:**
1. Plugin polls Graph API for new messages (or receives webhook)
2. New emails become incoming messages to the agent
3. Agent processes and responds

**Outbound flow:**
1. Agent uses `message` tool with `channel=microsoft365`
2. Plugin calls Graph API to send email
3. Email delivered from authenticated user's mailbox

## Troubleshooting

### "Need admin approval" during OAuth

Your organization requires admin consent. See Step 4.

### "Invalid client secret"

- Secret may have expired (check Azure Portal)
- Secret not saved correctly — run `openclaw secrets set MICROSOFT365_CLIENT_SECRET` again

### "Token refresh failed"

- Refresh token expired (rare, usually 90 days)
- Re-run `openclaw microsoft365 auth` to get new tokens

### "Insufficient privileges"

Missing API permissions. Verify all 5 permissions are added and admin consent granted.

## Security Notes

- Client secret is stored encrypted in OpenClaw secrets store
- Refresh token is stored in config (file permissions should be 600)
- Access tokens are short-lived and never persisted
- All API calls use HTTPS

## Files

```
extensions/microsoft365/
├── openclaw.plugin.json   # Plugin manifest
├── package.json           # Dependencies
├── index.ts              # Plugin entry point
└── src/
    ├── index.ts          # Exports
    ├── channel.ts        # Channel implementation
    ├── monitor.ts        # Email monitoring/polling
    ├── graph-client.ts   # Microsoft Graph API client
    └── types.ts          # TypeScript types
```
