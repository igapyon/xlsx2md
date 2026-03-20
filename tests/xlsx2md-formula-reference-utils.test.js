// @vitest-environment jsdom

import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";
import { loadModuleRegistry } from "./helpers/module-registry.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const formulaReferenceUtilsCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/formula-reference-utils.js"),
  "utf8"
);

function bootFormulaReferenceUtils() {
  document.body.innerHTML = "";
  loadModuleRegistry(__dirname);
  new Function(formulaReferenceUtilsCode)();
  return globalThis.__xlsx2mdModuleRegistry.getModule("formulaReferenceUtils");
}

describe("xlsx2md formula reference utils", () => {
  it("parses local and qualified cell references", () => {
    const module = bootFormulaReferenceUtils();
    const api = module.createFormulaReferenceUtilsApi({
      normalizeFormulaAddress: (address) => String(address || "").replace(/\$/g, "").toUpperCase()
    });

    expect(api.parseSimpleFormulaReference("=A1", "Sheet1")).toEqual({
      sheetName: "Sheet1",
      address: "A1"
    });
    expect(api.parseSimpleFormulaReference("='Sheet 2'!$B$3", "Sheet1")).toEqual({
      sheetName: "Sheet 2",
      address: "B3"
    });
  });

  it("normalizes sheet names and defined-name keys", () => {
    const module = bootFormulaReferenceUtils();
    const api = module.createFormulaReferenceUtilsApi({
      normalizeFormulaAddress: (address) => String(address || "")
    });

    expect(api.normalizeFormulaSheetName("'Sheet''A'")).toBe("Sheet'A");
    expect(api.normalizeDefinedNameKey(" total_value ")).toBe("TOTAL_VALUE");
  });

  it("parses sheet-scoped defined-name references but excludes cell refs", () => {
    const module = bootFormulaReferenceUtils();
    const api = module.createFormulaReferenceUtilsApi({
      normalizeFormulaAddress: (address) => String(address || "")
    });

    expect(api.parseSheetScopedDefinedNameReference("'Sheet 2'!Total_Value", "Sheet1")).toEqual({
      sheetName: "Sheet 2",
      name: "Total_Value"
    });
    expect(api.parseSheetScopedDefinedNameReference("Sheet2!B3", "Sheet1")).toBeNull();
  });
});
