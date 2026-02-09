import type { PluginRuntime } from "openclaw/plugin-sdk";

let runtime: PluginRuntime | null = null;

export function setMicrosoft365Runtime(next: PluginRuntime) {
  runtime = next;
}

export function getMicrosoft365Runtime(): PluginRuntime {
  if (!runtime) {
    throw new Error("Microsoft365 runtime not initialized");
  }
  return runtime;
}
