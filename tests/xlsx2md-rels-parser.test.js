// @vitest-environment jsdom

import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";
import { loadModuleRegistry, loadRuntimeEnv } from "./helpers/module-registry.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const xmlUtilsCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/xml-utils.js"),
  "utf8"
);
const relsParserCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/rels-parser.js"),
  "utf8"
);

function bootRelsParser() {
  document.body.innerHTML = "";
  loadModuleRegistry(__dirname);
  loadRuntimeEnv(__dirname);
  new Function(xmlUtilsCode)();
  new Function(relsParserCode)();
  return globalThis.__xlsx2mdModuleRegistry.getModule("relsParser").createRelsParserApi(
    globalThis.__xlsx2mdModuleRegistry.getModule("xmlUtils")
  );
}

describe("xlsx2md rels parser", () => {
  it("normalizes zip paths relative to the source path", () => {
    const api = bootRelsParser();

    expect(api.normalizeZipPath("xl/worksheets/sheet1.xml", "../drawings/drawing1.xml")).toBe("xl/drawings/drawing1.xml");
    expect(api.normalizeZipPath("xl/workbook.xml", "/docProps/app.xml")).toBe("docProps/app.xml");
  });

  it("builds rels paths next to the source file", () => {
    const api = bootRelsParser();

    expect(api.buildRelsPath("xl/workbook.xml")).toBe("xl/_rels/workbook.xml.rels");
    expect(api.buildRelsPath("xl/worksheets/sheet1.xml")).toBe("xl/worksheets/_rels/sheet1.xml.rels");
  });

  it("parses relationship targets into a map", () => {
    const api = bootRelsParser();
    const files = new Map([
      ["xl/_rels/workbook.xml.rels", new TextEncoder().encode(
        '<?xml version="1.0" encoding="UTF-8"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Target="worksheets/sheet1.xml"/>' +
        '<Relationship Id="rId2" Target="../sharedStrings.xml"/>' +
        "</Relationships>"
      )]
    ]);

    const rels = api.parseRelationships(files, "xl/_rels/workbook.xml.rels", "xl/workbook.xml");

    expect(rels.get("rId1")).toBe("xl/worksheets/sheet1.xml");
    expect(rels.get("rId2")).toBe("sharedStrings.xml");
  });
});
