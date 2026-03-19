// @vitest-environment jsdom

import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const formulaAstCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/formula-ast.js"),
  "utf8"
);

function bootFormulaAst() {
  document.body.innerHTML = "";
  new Function(formulaAstCode)();
  return globalThis.__xlsx2mdFormulaAst;
}

function createDeps() {
  return {
    normalizeFormulaAddress: (address) => String(address || "").replace(/\$/g, "").trim().toUpperCase(),
    parseSheetScopedDefinedNameReference: () => null,
    parseRangeAddress: (rawRange) => {
      const match = String(rawRange || "").match(/^([A-Z]+\d+):([A-Z]+\d+)$/i);
      return match ? { start: match[1].toUpperCase(), end: match[2].toUpperCase() } : null;
    },
    parseCellAddress: (address) => {
      const match = String(address || "").match(/^([A-Z]+)(\d+)$/i);
      if (!match) return { row: 0, col: 0 };
      let col = 0;
      for (const ch of match[1].toUpperCase()) {
        col = col * 26 + (ch.charCodeAt(0) - 64);
      }
      return { row: Number(match[2]), col };
    }
  };
}

describe("xlsx2md formula ast", () => {
  it("coerces booleans, numbers, and text scalars", () => {
    const module = bootFormulaAst();
    const api = module.createFormulaAstApi(createDeps());

    expect(api.coerceFormulaAstScalar("TRUE")).toBe(true);
    expect(api.coerceFormulaAstScalar("1,024")).toBe(1024);
    expect(api.coerceFormulaAstScalar("hello")).toBe("hello");
  });

  it("builds a scalar matrix from range entries", () => {
    const module = bootFormulaAst();
    const api = module.createFormulaAstApi(createDeps());

    const matrix = api.createFormulaAstRangeMatrix("Sheet1", "A1", "B2", () => ({
      rawValues: ["1", "TRUE", "x", "4"],
      numericValues: [1, 4]
    }));

    expect(matrix).toEqual([
      [1, true],
      ["x", 4]
    ]);
  });

  it("returns null for array AST results when serializing", () => {
    const module = bootFormulaAst();
    const api = module.createFormulaAstApi(createDeps());

    expect(api.serializeFormulaAstResult([1, 2])).toBeNull();
    expect(api.serializeFormulaAstResult(false)).toBe("FALSE");
  });
});
