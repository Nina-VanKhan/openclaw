/**
 * User secrets for agent tool use.
 *
 * Secrets are stored separately from auth profiles (which are for model providers).
 * These are user-defined secrets that agents can use in commands via env vars.
 *
 * Key principle: secret VALUES never enter model context. Only NAMES are exposed.
 * The agent writes commands using $SECRET_NAME, and the value is injected at exec time.
 */

export type SecretEntry = {
  /** The secret value (never logged, never sent to model). */
  value: string;
  /** Optional description shown to the agent. */
  description?: string;
  /** ISO timestamp when the secret was created. */
  createdAt: string;
  /** ISO timestamp when the secret was last updated. */
  updatedAt: string;
};

export type SecretsStore = {
  version: number;
  /** Map of secret name → secret entry. Names should be UPPER_SNAKE_CASE. */
  secrets: Record<string, SecretEntry>;
};

export type SecretMetadata = {
  name: string;
  description?: string;
  createdAt: string;
  updatedAt: string;
};

// Re-export SecretsConfig from config (single source of truth)
export type { SecretsConfig } from "../config/types.secrets.js";

export const SECRETS_STORE_VERSION = 1;
