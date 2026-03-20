// @vitest-environment jsdom

import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";
import { loadModuleRegistry } from "./helpers/module-registry.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const workbookLoaderCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/workbook-loader.js"),
  "utf8"
);

function bootWorkbookLoader() {
  document.body.innerHTML = "";
  loadModuleRegistry(__dirname);
  new Function(workbookLoaderCode)();
  return globalThis.__xlsx2mdModuleRegistry.getModule("workbookLoader");
}

describe("xlsx2md workbook loader", () => {
  it("parses defined names and skips _xlnm entries", () => {
    const api = bootWorkbookLoader();
    const workbookDoc = new DOMParser().parseFromString(
      `<?xml version="1.0" encoding="UTF-8"?>
      <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <definedNames>
          <definedName name="ValidName">Sheet1!$A$1</definedName>
          <definedName name="_xlnm.Print_Area">Sheet1!$A$1:$B$2</definedName>
          <definedName name="LocalName" localSheetId="1">Sheet2!$C$3</definedName>
        </definedNames>
      </workbook>`,
      "application/xml"
    );

    const result = api.parseDefinedNames(workbookDoc, ["Sheet1", "Sheet2"], (node) => node?.textContent || "");

    expect(result).toEqual([
      { name: "ValidName", formulaText: "=Sheet1!$A$1", localSheetName: null },
      { name: "LocalName", formulaText: "=Sheet2!$C$3", localSheetName: "Sheet2" }
    ]);
  });

  it("loads workbook parts and invokes worksheet parsing and post processing", async () => {
    const api = bootWorkbookLoader();
    const files = new Map([
      ["xl/workbook.xml", new TextEncoder().encode(
        `<?xml version="1.0" encoding="UTF-8"?>
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheets>
            <sheet name="First" r:id="rId1"/>
            <sheet name="Second" r:id="rId2"/>
          </sheets>
        </workbook>`
      )]
    ]);
    const seen = { parseWorksheet: [], post: 0 };

    const workbook = await api.parseWorkbook(new ArrayBuffer(0), "book.xlsx", {
      unzipEntries: async () => files,
      parseSharedStrings: () => ["A"],
      parseCellStyles: () => [{ borders: {}, numFmtId: 0, formatCode: "General" }],
      parseRelationships: () => new Map([["rId1", "xl/worksheets/sheet1.xml"], ["rId2", "xl/worksheets/sheet2.xml"]]),
      xmlToDocument: (xmlText) => new DOMParser().parseFromString(xmlText, "application/xml"),
      decodeXmlText: (bytes) => new TextDecoder().decode(bytes),
      getTextContent: (node) => node?.textContent || "",
      parseWorksheet: (_files, sheetName, sheetPath, sheetIndex) => {
        seen.parseWorksheet.push({ sheetName, sheetPath, sheetIndex });
        return { name: sheetName, index: sheetIndex, path: sheetPath, cells: [], merges: [], tables: [], images: [], charts: [], shapes: [], maxRow: 0, maxCol: 0 };
      },
      postProcessWorkbook: () => {
        seen.post += 1;
      }
    });

    expect(workbook.name).toBe("book.xlsx");
    expect(workbook.sharedStrings).toEqual(["A"]);
    expect(workbook.sheets).toHaveLength(2);
    expect(seen.parseWorksheet).toEqual([
      { sheetName: "First", sheetPath: "xl/worksheets/sheet1.xml", sheetIndex: 1 },
      { sheetName: "Second", sheetPath: "xl/worksheets/sheet2.xml", sheetIndex: 2 }
    ]);
    expect(seen.post).toBe(1);
  });

  it("throws when xl/workbook.xml is missing", async () => {
    const api = bootWorkbookLoader();

    await expect(api.parseWorkbook(new ArrayBuffer(0), "book.xlsx", {
      unzipEntries: async () => new Map(),
      parseSharedStrings: () => [],
      parseCellStyles: () => [],
      parseRelationships: () => new Map(),
      xmlToDocument: () => new DOMParser().parseFromString("<root/>", "application/xml"),
      decodeXmlText: () => "",
      getTextContent: () => "",
      parseWorksheet: () => ({})
    })).rejects.toThrow("xl/workbook.xml was not found.");
  });
});
