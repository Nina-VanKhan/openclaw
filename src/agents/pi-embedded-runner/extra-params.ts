import type { StreamFn } from "@mariozechner/pi-agent-core";
import type { SimpleStreamOptions } from "@mariozechner/pi-ai";
import { stream, streamSimple } from "@mariozechner/pi-ai";

import type { OpenClawConfig } from "../../config/config.js";
import { log } from "./logger.js";

/**
 * Resolve provider-specific extra params from model config.
 * Used to pass through stream params like temperature/maxTokens.
 *
 * @internal Exported for testing only
 */
export function resolveExtraParams(params: {
  cfg: OpenClawConfig | undefined;
  provider: string;
  modelId: string;
}): Record<string, unknown> | undefined {
  const modelKey = `${params.provider}/${params.modelId}`;
  const modelConfig = params.cfg?.agents?.defaults?.models?.[modelKey];
  return modelConfig?.params ? { ...modelConfig.params } : undefined;
}

type CacheControlTtl = "5m" | "1h";

function resolveCacheControlTtl(
  extraParams: Record<string, unknown> | undefined,
  provider: string,
  modelId: string,
): CacheControlTtl | undefined {
  const raw = extraParams?.cacheControlTtl;
  if (raw !== "5m" && raw !== "1h") {
    return undefined;
  }
  if (provider === "anthropic") {
    return raw;
  }
  if (provider === "openrouter" && modelId.startsWith("anthropic/")) {
    return raw;
  }
  return undefined;
}

function createStreamFnWithExtraParams(
  baseStreamFn: StreamFn | undefined,
  extraParams: Record<string, unknown> | undefined,
  provider: string,
  modelId: string,
): StreamFn | undefined {
  if (!extraParams || Object.keys(extraParams).length === 0) {
    return undefined;
  }

  const streamParams: Partial<SimpleStreamOptions> & {
    cacheControlTtl?: CacheControlTtl;
    toolChoice?: "auto" | "none" | "required";
  } = {};
  if (typeof extraParams.temperature === "number") {
    streamParams.temperature = extraParams.temperature;
  }
  if (typeof extraParams.maxTokens === "number") {
    streamParams.maxTokens = extraParams.maxTokens;
  }
  // Pass through toolChoice for OpenAI-compatible APIs (Ollama requires this for tool calls)
  if (
    extraParams.toolChoice === "auto" ||
    extraParams.toolChoice === "none" ||
    extraParams.toolChoice === "required"
  ) {
    streamParams.toolChoice = extraParams.toolChoice;
  }
  const cacheControlTtl = resolveCacheControlTtl(extraParams, provider, modelId);
  if (cacheControlTtl) {
    streamParams.cacheControlTtl = cacheControlTtl;
  }

  if (Object.keys(streamParams).length === 0) {
    return undefined;
  }

  log.debug(`creating streamFn wrapper with params: ${JSON.stringify(streamParams)}`);

  // Use `stream` (not `streamSimple`) when toolChoice is set, as streamSimple doesn't support it.
  // IMPORTANT: We must force `stream` when toolChoice is needed, even if baseStreamFn is set,
  // because streamSimple doesn't support toolChoice parameter.
  const needsFullStream = Boolean(streamParams.toolChoice);
  const underlying = needsFullStream ? stream : (baseStreamFn ?? streamSimple);

  const wrappedStreamFn: StreamFn = (model, context, options) => {
    // Filter out undefined values from options to prevent overriding streamParams
    const filteredOptions = options
      ? Object.fromEntries(Object.entries(options).filter(([, v]) => v !== undefined))
      : {};
    const merged = {
      ...streamParams,
      ...filteredOptions,
    };
    // Cast to StreamFn since `stream` has compatible runtime behavior but stricter types
    return (underlying as StreamFn)(model, context, merged);
  };

  return wrappedStreamFn;
}

/**
 * Apply extra params (like temperature) to an agent's streamFn.
 *
 * @internal Exported for testing
 */
export function applyExtraParamsToAgent(
  agent: { streamFn?: StreamFn },
  cfg: OpenClawConfig | undefined,
  provider: string,
  modelId: string,
  extraParamsOverride?: Record<string, unknown>,
): void {
  const extraParams = resolveExtraParams({
    cfg,
    provider,
    modelId,
  });
  const override =
    extraParamsOverride && Object.keys(extraParamsOverride).length > 0
      ? Object.fromEntries(
          Object.entries(extraParamsOverride).filter(([, value]) => value !== undefined),
        )
      : undefined;
  const merged = Object.assign({}, extraParams, override);
  const wrappedStreamFn = createStreamFnWithExtraParams(agent.streamFn, merged, provider, modelId);

  if (wrappedStreamFn) {
    log.debug(`applying extraParams to agent streamFn for ${provider}/${modelId}`);
    agent.streamFn = wrappedStreamFn;
  }
}
