// @vitest-environment jsdom

import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const cellFormatCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/cell-format.js"),
  "utf8"
);

function bootCellFormat() {
  document.body.innerHTML = "";
  new Function(cellFormatCode)();
  return globalThis.__xlsx2mdCellFormat;
}

describe("xlsx2md cell format", () => {
  it("formats date, percent, fraction, and currency values", () => {
    const api = bootCellFormat();

    expect(api.formatCellDisplayValue("13", { borders: {}, numFmtId: 14, formatCode: "yyyy/m/d" })).toBe("1900/1/12");
    expect(api.formatCellDisplayValue("0.987", { borders: {}, numFmtId: 10, formatCode: "0.0%" })).toBe("98.7%");
    expect(api.formatCellDisplayValue("0.75", { borders: {}, numFmtId: 12, formatCode: "# ?/?" })).toBe("3/4");
    expect(api.formatCellDisplayValue("1024768", { borders: {}, numFmtId: 42, formatCode: "¥ * #,##0" })).toBe("¥ 1,024,768");
  });

  it("formats special text and scientific notation patterns", () => {
    const api = bootCellFormat();

    expect(api.formatCellDisplayValue("1023456", { borders: {}, numFmtId: 186, formatCode: "[DBNum3]General" })).toBe("1 0 2 3 4 5 6");
    expect(api.formatCellDisplayValue("1023456", { borders: {}, numFmtId: 11, formatCode: "0.000000E+00" })).toBe("1.023456E+06");
    expect(api.formatCellDisplayValue("0", { borders: {}, numFmtId: 0, formatCode: '0;0;"-"' })).toBe("-");
  });

  it("parses date-like text into Excel serial-compatible numbers", () => {
    const api = bootCellFormat();

    expect(api.parseValueFunctionText("1,024")).toBe(1024);
    expect(api.parseValueFunctionText("2000-03-13")).toBe(36598);
    expect(api.parseValueFunctionText("3月14日")).toBe(36599);
    expect(api.parseValueFunctionText("")).toBeNull();
  });

  it("applies resolved formula formatting back onto a cell", () => {
    const api = bootCellFormat();
    const cell = {
      borders: { top: false, bottom: false, left: false, right: false },
      numFmtId: 10,
      formatCode: "0.0%",
      rawValue: "",
      outputValue: "",
      resolutionStatus: null,
      resolutionSource: null
    };

    api.applyResolvedFormulaValue(cell, "0.125", "legacy_resolver");

    expect(cell.rawValue).toBe("0.125");
    expect(cell.outputValue).toBe("12.5%");
    expect(cell.resolutionStatus).toBe("resolved");
    expect(cell.resolutionSource).toBe("legacy_resolver");
  });
});
