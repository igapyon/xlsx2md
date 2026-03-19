// @vitest-environment jsdom

import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const worksheetParserCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/worksheet-parser.js"),
  "utf8"
);

function bootWorksheetParser() {
  document.body.innerHTML = "";
  new Function(worksheetParserCode)();
  return globalThis.__xlsx2mdWorksheetParser;
}

function createDeps() {
  return {
    EMPTY_BORDERS: { top: false, bottom: false, left: false, right: false },
    xmlToDocument(xmlText) {
      return new DOMParser().parseFromString(xmlText, "application/xml");
    },
    decodeXmlText(bytes) {
      return new TextDecoder().decode(bytes);
    },
    getTextContent(node) {
      return (node?.textContent || "").replace(/\r\n/g, "\n");
    },
    parseCellAddress(address) {
      const match = String(address || "").match(/^([A-Z]+)(\d+)$/i);
      if (!match) return { row: 0, col: 0 };
      const letters = match[1].toUpperCase();
      let col = 0;
      for (const ch of letters) col = col * 26 + (ch.charCodeAt(0) - 64);
      return { row: Number(match[2]), col };
    },
    parseRangeRef(ref) {
      const [start, end] = String(ref).split(":");
      const s = this.parseCellAddress(start);
      const e = this.parseCellAddress(end || start);
      return { startRow: s.row, startCol: s.col, endRow: e.row, endCol: e.col, ref };
    },
    parseWorksheetTables() {
      return [];
    },
    parseDrawingImages() {
      return [];
    },
    parseDrawingCharts() {
      return [];
    },
    parseDrawingShapes() {
      return [];
    },
    formatCellDisplayValue(rawValue, style) {
      if (style.numFmtId === 10) {
        return `${(Number(rawValue) * 100).toFixed(1)}%`;
      }
      return null;
    },
    buildAssetDeps() {
      return {};
    },
    lettersToCol(letters) {
      let result = 0;
      for (const ch of String(letters).toUpperCase()) result = result * 26 + (ch.charCodeAt(0) - 64);
      return result;
    },
    colToLetters(col) {
      let current = col;
      let result = "";
      while (current > 0) {
        const remainder = (current - 1) % 26;
        result = String.fromCharCode(65 + remainder) + result;
        current = Math.floor((current - 1) / 26);
      }
      return result;
    }
  };
}

describe("xlsx2md worksheet parser", () => {
  it("extracts shared string and boolean cell values", () => {
    const api = bootWorksheetParser();
    const cellStyle = { borders: {}, numFmtId: 0, formatCode: "General" };
    const sharedCell = new DOMParser().parseFromString('<c t="s"><v>1</v></c>', "application/xml").documentElement;
    const boolCell = new DOMParser().parseFromString('<c t="b"><v>1</v></c>', "application/xml").documentElement;
    const deps = createDeps();

    expect(api.extractCellOutputValue(sharedCell, ["A", "B"], cellStyle, deps)).toMatchObject({
      valueType: "s",
      outputValue: "B"
    });
    expect(api.extractCellOutputValue(boolCell, [], cellStyle, deps)).toMatchObject({
      valueType: "b",
      outputValue: "TRUE"
    });
  });

  it("translates shared formulas across relative references", () => {
    const api = bootWorksheetParser();
    const deps = createDeps();

    const translated = api.translateSharedFormula("=A1+$B$1", "C1", "C3", deps);

    expect(translated).toBe("=A3+$B$1");
  });

  it("parses worksheet cells, merges, and shared formulas", () => {
    const api = bootWorksheetParser();
    const deps = createDeps();
    const worksheetXml = `<?xml version="1.0" encoding="UTF-8"?>
      <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheetData>
          <row r="1">
            <c r="A1" t="s"><v>0</v></c>
            <c r="B1" s="1"><v>0.125</v></c>
            <c r="C1"><f t="shared" si="0">A1</f><v>1</v></c>
          </row>
          <row r="2">
            <c r="C2"><f t="shared" si="0"/><v>2</v></c>
          </row>
        </sheetData>
        <mergeCells count="1">
          <mergeCell ref="A1:B2"/>
        </mergeCells>
      </worksheet>`;
    const files = new Map([
      ["xl/worksheets/sheet1.xml", new TextEncoder().encode(worksheetXml)]
    ]);
    const cellStyles = [
      { borders: { top: false, bottom: false, left: false, right: false }, numFmtId: 0, formatCode: "General" },
      { borders: { top: false, bottom: false, left: false, right: false }, numFmtId: 10, formatCode: "0.0%" }
    ];

    const sheet = api.parseWorksheet(files, "Sheet1", "xl/worksheets/sheet1.xml", 1, ["Hello"], cellStyles, deps);

    expect(sheet.cells.find((cell) => cell.address === "A1")?.outputValue).toBe("Hello");
    expect(sheet.cells.find((cell) => cell.address === "B1")?.outputValue).toBe("12.5%");
    expect(sheet.cells.find((cell) => cell.address === "C1")?.formulaText).toBe("=A1");
    expect(sheet.cells.find((cell) => cell.address === "C2")?.formulaText).toBe("=A2");
    expect(sheet.merges).toEqual([{ startRow: 1, startCol: 1, endRow: 2, endCol: 2, ref: "A1:B2" }]);
    expect(sheet.maxRow).toBe(2);
    expect(sheet.maxCol).toBe(3);
  });
});
