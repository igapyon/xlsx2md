// @vitest-environment jsdom

import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const formulaLegacyCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/formula-legacy.js"),
  "utf8"
);

function bootFormulaLegacy() {
  document.body.innerHTML = "";
  new Function(formulaLegacyCode)();
  return globalThis.__xlsx2mdFormulaLegacy;
}

function createDeps(overrides = {}) {
  const parseCellAddress = (address) => {
    const match = String(address || "").match(/^([A-Z]+)(\d+)$/i);
    if (!match) return { row: 0, col: 0 };
    let col = 0;
    for (const ch of match[1].toUpperCase()) {
      col = col * 26 + (ch.charCodeAt(0) - 64);
    }
    return { row: Number(match[2]), col };
  };
  return {
    normalizeFormulaSheetName: (rawName) => String(rawName || "").replace(/^'/, "").replace(/'$/, "").replace(/''/g, "'"),
    normalizeFormulaAddress: (address) => String(address || "").replace(/\$/g, "").trim().toUpperCase(),
    parseSimpleFormulaReference: (formulaText, currentSheetName) => {
      const normalized = String(formulaText || "").trim().replace(/^=/, "");
      const localMatch = normalized.match(/^([A-Z]+\d+)$/i);
      if (localMatch) return { sheetName: currentSheetName, address: localMatch[1].toUpperCase() };
      return null;
    },
    parseSheetScopedDefinedNameReference: () => null,
    parseRangeAddress: (rawRange) => {
      const match = String(rawRange || "").match(/^([A-Z]+\d+):([A-Z]+\d+)$/i);
      return match ? { start: match[1].toUpperCase(), end: match[2].toUpperCase() } : null;
    },
    parseCellAddress,
    colToLetters: (col) => {
      let current = col;
      let result = "";
      while (current > 0) {
        const remainder = (current - 1) % 26;
        result = String.fromCharCode(65 + remainder) + result;
        current = Math.floor((current - 1) / 26);
      }
      return result;
    },
    tryResolveFormulaExpression: () => null,
    getDefinedNameScalarValue: () => null,
    getDefinedNameRangeRef: () => null,
    getStructuredRangeRef: () => null,
    cellFormat: {
      formatTextFunctionValue: (value, formatText) => `${formatText}:${value}`,
      parseValueFunctionText: (source) => Number(source),
      datePartsToExcelSerial: (yyyy, mm, dd) => Number(`${yyyy}${String(mm).padStart(2, "0")}${String(dd).padStart(2, "0")}`),
      parseDateLikeParts: (source) => {
        const match = String(source || "").match(/^(\d{4})-(\d{2})-(\d{2})$/);
        return match ? { yyyy: match[1], mm: match[2], dd: match[3] } : null;
      }
    },
    ...overrides
  };
}

describe("xlsx2md formula legacy", () => {
  it("splits function arguments with nesting and quoted commas", () => {
    const module = bootFormulaLegacy();
    const api = module.createFormulaLegacyApi(createDeps());

    expect(api.splitFormulaArguments('A1,"x,y",SUM(B1:B2)')).toEqual([
      "A1",
      '"x,y"',
      "SUM(B1:B2)"
    ]);
  });

  it("resolves structured ranges through the helper API", () => {
    const module = bootFormulaLegacy();
    const api = module.createFormulaLegacyApi(createDeps({
      getStructuredRangeRef: () => (sheetName, text) => {
        if (sheetName === "Sheet1" && text === "Sales[Amount]") {
          return { sheetName: "Sheet1", start: "D3", end: "D4" };
        }
        return null;
      }
    }));

    expect(api.parseQualifiedRangeReference("Sales[Amount]", "Sheet1")).toEqual({
      sheetName: "Sheet1",
      start: "D3",
      end: "D4"
    });
  });

  it("evaluates legacy IF expressions without the AST evaluator", () => {
    const module = bootFormulaLegacy();
    const api = module.createFormulaLegacyApi(createDeps());

    const resolved = api.tryResolveFormulaExpressionLegacy(
      'IF(1=1,"ok","ng")',
      "Sheet1",
      () => ""
    );

    expect(resolved).toBe("ok");
  });
});
