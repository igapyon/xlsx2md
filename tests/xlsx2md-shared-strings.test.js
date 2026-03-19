// @vitest-environment jsdom

import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const sharedStringsCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/shared-strings.js"),
  "utf8"
);

function bootSharedStrings() {
  document.body.innerHTML = "";
  new Function(sharedStringsCode)();
  return globalThis.__xlsx2mdSharedStrings;
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

    expect(api.parseSharedStrings(files)).toEqual(["通常のテキスト", "分割テキスト"]);
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

    expect(api.parseSharedStrings(files)).toEqual(["line1\nline2"]);
  });
});
