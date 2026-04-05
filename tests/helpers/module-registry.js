import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import path from "node:path";

export function loadModuleRegistry(testDir) {
  globalThis.__xlsx2mdNodeRequire ??= createRequire(import.meta.url);
  const moduleRegistryCode = readFileSync(
    path.resolve(testDir, "../src/js/module-registry.js"),
    "utf8"
  );
  const moduleRegistryAccessCode = readFileSync(
    path.resolve(testDir, "../src/js/module-registry-access.js"),
    "utf8"
  );
  new Function(moduleRegistryCode)();
  new Function(moduleRegistryAccessCode)();
  return globalThis.__xlsx2mdModuleRegistry;
}

export function loadRuntimeEnv(testDir) {
  const runtimeEnvCode = readFileSync(
    path.resolve(testDir, "../src/js/runtime-env.js"),
    "utf8"
  );
  new Function(runtimeEnvCode)();
  return globalThis.__xlsx2mdModuleRegistry?.getModule("runtimeEnv") || null;
}
