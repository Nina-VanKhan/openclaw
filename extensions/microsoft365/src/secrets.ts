/**
 * Secrets helper for Microsoft 365 plugin.
 *
 * Reads secrets from OpenClaw's secrets store (~/.openclaw/secrets.json).
 * This is a standalone implementation so the plugin works regardless of
 * whether the secrets module is merged upstream.
 */

import * as fs from "node:fs";
import * as path from "node:path";
import * as os from "node:os";

type SecretEntry = {
  value: string;
  description?: string;
  createdAt: string;
  updatedAt: string;
};

type SecretsStore = {
  version: number;
  secrets: Record<string, SecretEntry>;
};

/**
 * Resolve the path to OpenClaw's secrets file.
 */
function resolveSecretsPath(): string {
  // Check for XDG config, fall back to ~/.openclaw
  const xdgState = process.env.XDG_STATE_HOME;
  const baseDir = xdgState
    ? path.join(xdgState, "openclaw")
    : path.join(os.homedir(), ".openclaw");
  return path.join(baseDir, "secrets.json");
}

/**
 * Load the secrets store from disk.
 */
function loadSecretsStore(): SecretsStore | null {
  const filePath = resolveSecretsPath();
  try {
    if (!fs.existsSync(filePath)) {
      return null;
    }
    const raw = fs.readFileSync(filePath, "utf8");
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== "object" || !parsed.secrets) {
      return null;
    }
    return parsed as SecretsStore;
  } catch {
    return null;
  }
}

/**
 * Get a secret value by name.
 *
 * @param name - The secret name (e.g., "MICROSOFT365_CLIENT_SECRET")
 * @returns The secret value, or undefined if not found
 */
export function getSecret(name: string): string | undefined {
  const store = loadSecretsStore();
  if (!store) {
    return undefined;
  }
  return store.secrets[name]?.value;
}

/**
 * Check if a secret exists.
 */
export function hasSecret(name: string): boolean {
  return getSecret(name) !== undefined;
}

/**
 * Secret name constants for this plugin.
 */
export const SECRETS = {
  CLIENT_SECRET: "MICROSOFT365_CLIENT_SECRET",
} as const;
