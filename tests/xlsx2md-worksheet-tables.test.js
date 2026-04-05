// @vitest-environment jsdom

import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";
import { loadModuleRegistry, loadRuntimeEnv } from "./helpers/module-registry.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const worksheetTablesCode = readFileSync(
  path.resolve(__dirname, "../src/js/worksheet-tables.js"),
  "utf8"
);

function bootWorksheetTables() {
  document.body.innerHTML = "";
  loadModuleRegistry(__dirname);
  loadRuntimeEnv(__dirname);
  new Function(worksheetTablesCode)();
  return globalThis.__xlsx2mdModuleRegistry.getModule("worksheetTables");
}

describe("xlsx2md worksheet tables", () => {
  it("normalizes structured table keys", () => {
    const api = bootWorksheetTables();

    expect(api.normalizeStructuredTableKey("  Ｔａｂｌｅ １ ")).toBe("TABLE 1");
  });

  it("returns no tables when no table parts exist", () => {
    const api = bootWorksheetTables();
    const worksheetDoc = new DOMParser().parseFromString(
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></worksheet>',
      "application/xml"
    );

    expect(api.parseWorksheetTables(new Map(), worksheetDoc, "Sheet1", "xl/worksheets/sheet1.xml")).toEqual([]);
  });

  it("resolves worksheet table parts through relationships", () => {
    const api = bootWorksheetTables();
    const worksheetXml = `<?xml version="1.0" encoding="UTF-8"?>
      <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <tableParts count="1">
          <tablePart r:id="rId1"/>
        </tableParts>
      </worksheet>`;
    const relsXml = `<?xml version="1.0" encoding="UTF-8"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Target="../tables/table1.xml"/>
      </Relationships>`;
    const tableXml = `<?xml version="1.0" encoding="UTF-8"?>
      <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
             id="1"
             name="SalesTable"
             displayName="Sales Table"
             ref="$B$2:$D$5"
             headerRowCount="1"
             totalsRowCount="1">
        <tableColumns count="3">
          <tableColumn id="1" name="Code"/>
          <tableColumn id="2" name="Name"/>
          <tableColumn id="3" name="Amount"/>
        </tableColumns>
      </table>`;
    const files = new Map([
      ["xl/worksheets/_rels/sheet1.xml.rels", new TextEncoder().encode(relsXml)],
      ["xl/tables/table1.xml", new TextEncoder().encode(tableXml)]
    ]);
    const worksheetDoc = new DOMParser().parseFromString(worksheetXml, "application/xml");

    const tables = api.parseWorksheetTables(files, worksheetDoc, "Report", "xl/worksheets/sheet1.xml");

    expect(tables).toEqual([
      {
        sheetName: "Report",
        name: "SalesTable",
        displayName: "Sales Table",
        start: "B2",
        end: "D5",
        columns: ["Code", "Name", "Amount"],
        headerRowCount: 1,
        totalsRowCount: 1
      }
    ]);
  });

  it("ignores table definitions with invalid ranges", () => {
    const api = bootWorksheetTables();
    const worksheetXml = `<?xml version="1.0" encoding="UTF-8"?>
      <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <tableParts count="1">
          <tablePart r:id="rId1"/>
        </tableParts>
      </worksheet>`;
    const relsXml = `<?xml version="1.0" encoding="UTF-8"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Target="../tables/table1.xml"/>
      </Relationships>`;
    const tableXml = `<?xml version="1.0" encoding="UTF-8"?>
      <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
             id="1"
             name="BrokenTable"
             ref="not-a-range"/>`;
    const files = new Map([
      ["xl/worksheets/_rels/sheet1.xml.rels", new TextEncoder().encode(relsXml)],
      ["xl/tables/table1.xml", new TextEncoder().encode(tableXml)]
    ]);
    const worksheetDoc = new DOMParser().parseFromString(worksheetXml, "application/xml");

    expect(api.parseWorksheetTables(files, worksheetDoc, "Report", "xl/worksheets/sheet1.xml")).toEqual([]);
  });
});
