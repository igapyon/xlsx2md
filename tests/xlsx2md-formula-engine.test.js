// @vitest-environment jsdom

import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";
import { loadModuleRegistry } from "./helpers/module-registry.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const formulaEngineCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/formula-engine.js"),
  "utf8"
);

function bootFormulaEngine() {
  document.body.innerHTML = "";
  loadModuleRegistry(__dirname);
  new Function(formulaEngineCode)();
  return globalThis.__xlsx2mdModuleRegistry.getModule("formulaEngine");
}

describe("xlsx2md formula engine", () => {
  it("prefers direct defined-name values before other resolvers", () => {
    const module = bootFormulaEngine();
    const api = module.createFormulaEngineApi({
      getDefinedNameScalarValue: () => (sheetName, name) => (sheetName === "Sheet1" && name === "Total" ? "42" : null),
      tryResolveFormulaExpressionWithAst: () => "999",
      tryResolveFormulaExpressionLegacy: () => "888"
    });

    expect(api.tryResolveFormulaExpressionDetailed("=Total", "Sheet1", () => "")).toEqual({
      value: "42",
      source: "legacy_resolver"
    });
  });

  it("uses the ast evaluator before the legacy resolver", () => {
    const module = bootFormulaEngine();
    const api = module.createFormulaEngineApi({
      getDefinedNameScalarValue: () => null,
      tryResolveFormulaExpressionWithAst: () => "12",
      tryResolveFormulaExpressionLegacy: () => "34"
    });

    expect(api.tryResolveFormulaExpressionDetailed("=A1+B1", "Sheet1", () => "")).toEqual({
      value: "12",
      source: "ast_evaluator"
    });
  });

  it("falls back to the legacy resolver and exposes value-only helper", () => {
    const module = bootFormulaEngine();
    const api = module.createFormulaEngineApi({
      getDefinedNameScalarValue: () => null,
      tryResolveFormulaExpressionWithAst: () => null,
      tryResolveFormulaExpressionLegacy: () => "56"
    });

    expect(api.tryResolveFormulaExpressionDetailed("=A1+B1", "Sheet1", () => "")).toEqual({
      value: "56",
      source: "legacy_resolver"
    });
    expect(api.tryResolveFormulaExpression("=A1+B1", "Sheet1", () => "")).toBe("56");
  });
});
