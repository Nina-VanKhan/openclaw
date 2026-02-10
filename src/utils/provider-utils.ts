/**
 * Utility functions for provider-specific logic and capabilities.
 */

/**
 * Returns true if the provider requires reasoning to be wrapped in tags
 * (e.g. <think> and <final>) in the text stream, rather than using native
 * API fields for reasoning/thinking.
 */
export function isReasoningTagProvider(provider: string | undefined | null): boolean {
  if (!provider) {
    return false;
  }
  const normalized = provider.trim().toLowerCase();

  // Check for exact matches or known prefixes/substrings for reasoning providers
  // NOTE: Ollama was removed from this list because:
  // 1. Ollama models with native reasoning use <think> tags internally
  // 2. Adding <final> wrapper instructions causes <final> tags to leak into UI
  //    (dashboard server.js only strips <think>, not <final>)
  // 3. Ollama models work correctly without the <final> tag requirement
  if (normalized === "google-gemini-cli" || normalized === "google-generative-ai") {
    return true;
  }

  // Handle google-antigravity and its model variations (e.g. google-antigravity/gemini-3)
  if (normalized.includes("google-antigravity")) {
    return true;
  }

  // Handle Minimax (M2.1 is chatty/reasoning-like)
  if (normalized.includes("minimax")) {
    return true;
  }

  return false;
}
