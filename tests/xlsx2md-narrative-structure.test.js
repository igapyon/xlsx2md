// @vitest-environment jsdom

import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";
import { loadModuleRegistry } from "./helpers/module-registry.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const markdownNormalizeCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/markdown-normalize.js"),
  "utf8"
);
const narrativeStructureCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/narrative-structure.js"),
  "utf8"
);

function bootNarrativeStructure() {
  document.body.innerHTML = "";
  loadModuleRegistry(__dirname);
  new Function(markdownNormalizeCode)();
  new Function(narrativeStructureCode)();
  return globalThis.__xlsx2mdModuleRegistry.getModule("narrativeStructure");
}

describe("xlsx2md narrative structure", () => {
  it("renders an indented parent-child block as heading plus bullets", () => {
    const api = bootNarrativeStructure();
    const markdown = api.renderNarrativeBlock({
      startRow: 1,
      startCol: 1,
      endRow: 3,
      lines: ["通常のテキスト", "字下げされたテキスト", "テキスト"],
      items: [
        { row: 1, startCol: 1, text: "通常のテキスト", cellValues: ["通常のテキスト"] },
        { row: 2, startCol: 2, text: "字下げされたテキスト", cellValues: ["字下げされたテキスト"] },
        { row: 3, startCol: 2, text: "テキスト", cellValues: ["テキスト"] }
      ]
    });

    expect(markdown).toContain("### 通常のテキスト");
    expect(markdown).toContain("- 字下げされたテキスト");
    expect(markdown).toContain("- テキスト");
  });

  it("renders flat narrative rows as plain paragraphs", () => {
    const api = bootNarrativeStructure();
    const markdown = api.renderNarrativeBlock({
      startRow: 1,
      startCol: 1,
      endRow: 2,
      lines: ["一行目", "二行目"],
      items: [
        { row: 1, startCol: 1, text: "一行目", cellValues: ["一行目"] },
        { row: 2, startCol: 1, text: "二行目", cellValues: ["二行目"] }
      ]
    });

    expect(markdown).toBe("一行目\n\n二行目");
  });

  it("starts a new heading when indentation returns to the parent level", () => {
    const api = bootNarrativeStructure();
    const markdown = api.renderNarrativeBlock({
      startRow: 1,
      startCol: 1,
      endRow: 5,
      lines: ["親1", "子1", "子2", "親2", "子3"],
      items: [
        { row: 1, startCol: 1, text: "親1", cellValues: ["親1"] },
        { row: 2, startCol: 2, text: "子1", cellValues: ["子1"] },
        { row: 3, startCol: 2, text: "子2", cellValues: ["子2"] },
        { row: 4, startCol: 1, text: "親2", cellValues: ["親2"] },
        { row: 5, startCol: 3, text: "子3", cellValues: ["子3"] }
      ]
    });

    expect(markdown).toContain("### 親1");
    expect(markdown).toContain("- 子1");
    expect(markdown).toContain("- 子2");
    expect(markdown).toContain("### 親2");
    expect(markdown).toContain("- 子3");
  });

  it("detects a heading block only when the second item is indented deeper", () => {
    const api = bootNarrativeStructure();

    expect(api.isSectionHeadingNarrativeBlock({
      startRow: 1,
      startCol: 1,
      endRow: 2,
      lines: ["親", "子"],
      items: [
        { row: 1, startCol: 1, text: "親", cellValues: ["親"] },
        { row: 2, startCol: 2, text: "子", cellValues: ["子"] }
      ]
    })).toBe(true);

    expect(api.isSectionHeadingNarrativeBlock({
      startRow: 1,
      startCol: 1,
      endRow: 2,
      lines: ["一行目", "二行目"],
      items: [
        { row: 1, startCol: 1, text: "一行目", cellValues: ["一行目"] },
        { row: 2, startCol: 1, text: "二行目", cellValues: ["二行目"] }
      ]
    })).toBe(false);
  });
});
