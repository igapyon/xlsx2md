// @vitest-environment jsdom

import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";
import { loadModuleRegistry, loadRuntimeEnv } from "./helpers/module-registry.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const stylesParserCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/styles-parser.js"),
  "utf8"
);

function bootStylesParser() {
  document.body.innerHTML = "";
  loadModuleRegistry(__dirname);
  loadRuntimeEnv(__dirname);
  new Function(stylesParserCode)();
  return globalThis.__xlsx2mdModuleRegistry.getModule("stylesParser");
}

describe("xlsx2md styles parser", () => {
  it("returns a general default style when styles.xml is missing", () => {
    const api = bootStylesParser();

    expect(api.parseCellStyles(new Map())).toEqual([
      {
        borders: { top: false, bottom: false, left: false, right: false },
        numFmtId: 0,
        formatCode: "General"
      }
    ]);
  });

  it("detects border sides from style attributes or child nodes", () => {
    const api = bootStylesParser();
    const doc = new DOMParser().parseFromString(
      "<root><top style=\"thin\"/><bottom><color rgb=\"000000\"/></bottom><left/><right/></root>",
      "application/xml"
    );

    expect(api.hasBorderSide(doc.getElementsByTagName("top")[0])).toBe(true);
    expect(api.hasBorderSide(doc.getElementsByTagName("bottom")[0])).toBe(true);
    expect(api.hasBorderSide(doc.getElementsByTagName("left")[0])).toBe(false);
    expect(api.hasBorderSide(doc.getElementsByTagName("right")[0])).toBe(false);
  });

  it("parses borders and custom number formats from styles.xml", () => {
    const api = bootStylesParser();
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
      <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <numFmts count="1">
          <numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>
        </numFmts>
        <borders count="2">
          <border><left/><right/><top/><bottom/></border>
          <border><left style="thin"/><right/><top style="thin"/><bottom/></border>
        </borders>
        <cellXfs count="2">
          <xf numFmtId="0" borderId="0"/>
          <xf numFmtId="164" borderId="1"/>
        </cellXfs>
      </styleSheet>`;
    const files = new Map([
      ["xl/styles.xml", new TextEncoder().encode(xml)]
    ]);

    const styles = api.parseCellStyles(files);

    expect(styles).toHaveLength(2);
    expect(styles[0]).toEqual({
      borders: { top: false, bottom: false, left: false, right: false },
      numFmtId: 0,
      formatCode: "General"
    });
    expect(styles[1]).toEqual({
      borders: { top: true, bottom: false, left: true, right: false },
      numFmtId: 164,
      formatCode: "yyyy-mm-dd"
    });
  });

  it("falls back to built-in format codes when no custom numFmt exists", () => {
    const api = bootStylesParser();
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
      <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <borders count="1">
          <border><left/><right/><top/><bottom/></border>
        </borders>
        <cellXfs count="1">
          <xf numFmtId="14" borderId="0"/>
        </cellXfs>
      </styleSheet>`;
    const files = new Map([
      ["xl/styles.xml", new TextEncoder().encode(xml)]
    ]);

    const styles = api.parseCellStyles(files);

    expect(styles[0].formatCode).toBe("yyyy/m/d");
  });
});
