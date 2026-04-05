// @vitest-environment jsdom

import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";
import { loadModuleRegistry, loadRuntimeEnv } from "./helpers/module-registry.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const sharedStringsCode = readFileSync(
  path.resolve(__dirname, "../src/js/shared-strings.js"),
  "utf8"
);

function bootSharedStrings() {
  document.body.innerHTML = "";
  loadModuleRegistry(__dirname);
  loadRuntimeEnv(__dirname);
  new Function(sharedStringsCode)();
  return globalThis.__xlsx2mdModuleRegistry.getModule("sharedStrings");
}

describe("xlsx2md shared strings", () => {
  it("returns an empty list when sharedStrings.xml is missing", () => {
    const api = bootSharedStrings();

    expect(api.parseSharedStrings(new Map())).toEqual([]);
  });

  it("collects plain and rich text runs", () => {
    const api = bootSharedStrings();
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
      <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
        <si><t>通常のテキスト</t></si>
        <si><r><t>分割</t></r><r><t>テキスト</t></r></si>
      </sst>`;
    const files = new Map([
      ["xl/sharedStrings.xml", new TextEncoder().encode(xml)]
    ]);

    expect(api.parseSharedStrings(files)).toEqual([
      { text: "通常のテキスト", runs: null },
      { text: "分割テキスト", runs: null }
    ]);
  });

  it("skips phonetic text nodes and normalizes CRLF to LF", () => {
    const api = bootSharedStrings();
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
      <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
        <si>
          <t>line1&#13;&#10;line2</t>
          <rPh sb="0" eb="1"><t>phonetic</t></rPh>
          <phoneticPr fontId="1"/>
        </si>
      </sst>`;
    const files = new Map([
      ["xl/sharedStrings.xml", new TextEncoder().encode(xml)]
    ]);

    expect(api.parseSharedStrings(files)).toEqual([{ text: "line1\nline2", runs: null }]);
  });

  it("preserves supported rich text emphasis flags", () => {
    const api = bootSharedStrings();
    const xml = `<?xml version="1.0" encoding="UTF-8"?>
      <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
        <si>
          <r><rPr><b/></rPr><t>Bold</t></r>
          <r><rPr><i/></rPr><t>Italic</t></r>
          <r><rPr><strike/></rPr><t>Strike</t></r>
          <r><rPr><u/></rPr><t>Under</t></r>
        </si>
      </sst>`;
    const files = new Map([
      ["xl/sharedStrings.xml", new TextEncoder().encode(xml)]
    ]);

    expect(api.parseSharedStrings(files)).toEqual([{
      text: "BoldItalicStrikeUnder",
      runs: [
        { text: "Bold", bold: true, italic: false, strike: false, underline: false },
        { text: "Italic", bold: false, italic: true, strike: false, underline: false },
        { text: "Strike", bold: false, italic: false, strike: true, underline: false },
        { text: "Under", bold: false, italic: false, strike: false, underline: true }
      ]
    }]);
  });
});
