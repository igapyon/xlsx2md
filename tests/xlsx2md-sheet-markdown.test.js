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
const markdownEscapeCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/markdown-escape.js"),
  "utf8"
);
const markdownTableEscapeCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/markdown-table-escape.js"),
  "utf8"
);
const richTextParserCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/rich-text-parser.js"),
  "utf8"
);
const richTextPlainFormatterCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/rich-text-plain-formatter.js"),
  "utf8"
);
const richTextGithubFormatterCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/rich-text-github-formatter.js"),
  "utf8"
);
const richTextRendererCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/rich-text-renderer.js"),
  "utf8"
);
const sheetMarkdownCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/sheet-markdown.js"),
  "utf8"
);

function bootSheetMarkdown() {
  document.body.innerHTML = "";
  loadModuleRegistry(__dirname);
  new Function(markdownNormalizeCode)();
  new Function(markdownEscapeCode)();
  new Function(markdownTableEscapeCode)();
  new Function(richTextParserCode)();
  new Function(richTextPlainFormatterCode)();
  new Function(richTextGithubFormatterCode)();
  new Function(richTextRendererCode)();
  new Function(sheetMarkdownCode)();
  return globalThis.__xlsx2mdModuleRegistry.getModule("sheetMarkdown");
}

function createDeps(overrides = {}) {
  return {
    renderNarrativeBlock: (block) => `### ${block.lines[0]}`,
    detectTableCandidates: () => [],
    matrixFromCandidate: () => [],
    renderMarkdownTable: (rows) => rows.map((row) => `| ${row.join(" | ")} |`).join("\n"),
    createOutputFileName: (_workbookName, sheetIndex, sheetName) => `${sheetIndex}_${sheetName}.md`,
    extractShapeBlocks: () => [],
    renderHierarchicalRawEntries: () => [],
    parseCellAddress: (address) => {
      const match = String(address || "").match(/^([A-Z]+)(\d+)$/i);
      if (!match) return { row: 0, col: 0 };
      let col = 0;
      for (const ch of match[1].toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
      return { row: Number(match[2]), col };
    },
    formatRange: (startRow, startCol, endRow, endCol) => `${startRow}:${startCol}-${endRow}:${endCol}`,
    colToLetters: (col) => String.fromCharCode(64 + col),
    normalizeMarkdownText: (text) => String(text || "").replace(/\r\n?|\n/g, " ").replace(/\s+/g, " ").trim(),
    defaultCellWidthEmu: 1,
    defaultCellHeightEmu: 1,
    shapeBlockGapXEmu: 1,
    shapeBlockGapYEmu: 1,
    ...overrides
  };
}

describe("xlsx2md sheet markdown", () => {
  it("builds cell maps and narrative row segments", () => {
    const module = bootSheetMarkdown();
    const api = module.createSheetMarkdownApi(createDeps());
    const sheet = {
      cells: [
        { row: 1, col: 1, outputValue: "A", rawValue: "A" },
        { row: 1, col: 2, outputValue: "B", rawValue: "B" },
        { row: 1, col: 8, outputValue: "C", rawValue: "C" }
      ]
    };

    expect(api.buildCellMap(sheet).get("1:2")?.outputValue).toBe("B");
    expect(api.splitNarrativeRowSegments(sheet.cells, {})).toEqual([
      { startCol: 1, values: ["A", "B"] },
      { startCol: 8, values: ["C"] }
    ]);
  });

  it("extracts narrative blocks outside detected tables", () => {
    const module = bootSheetMarkdown();
    const api = module.createSheetMarkdownApi(createDeps());
    const sheet = {
      cells: [
        { row: 1, col: 1, outputValue: "Heading", rawValue: "Heading" },
        { row: 2, col: 2, outputValue: "Detail", rawValue: "Detail" },
        { row: 5, col: 1, outputValue: "TableCell", rawValue: "TableCell" }
      ],
      charts: [],
      images: []
    };
    const tables = [{ startRow: 5, startCol: 1, endRow: 5, endCol: 1 }];

    expect(api.extractNarrativeBlocks(sheet, tables, {})).toHaveLength(1);
    expect(api.extractNarrativeBlocks(sheet, tables, {})[0].lines).toEqual(["Heading", "Detail"]);
  });

  it("converts a minimal sheet into markdown", () => {
    const module = bootSheetMarkdown();
    const api = module.createSheetMarkdownApi(createDeps({
      renderNarrativeBlock: (block) => block.lines.join("\n")
    }));
    const workbook = { name: "book.xlsx", sheets: [] };
    const sheet = {
      name: "Sheet1",
      index: 1,
      cells: [
        { address: "A1", row: 1, col: 1, outputValue: "Hello", rawValue: "Hello", formulaText: "", resolutionStatus: null, resolutionSource: null },
        { address: "B2", row: 2, col: 2, outputValue: "World", rawValue: "World", formulaText: "", resolutionStatus: null, resolutionSource: null }
      ],
      merges: [],
      images: [],
      charts: [],
      shapes: []
    };

    const result = api.convertSheetToMarkdown(workbook, sheet, {});

    expect(result.fileName).toBe("1_Sheet1.md");
    expect(result.markdown).toContain("# Sheet1");
    expect(result.markdown).toContain("Workbook: book.xlsx");
    expect(result.summary.narrativeBlocks).toBe(1);
  });

  it("normalizes cell line breaks into spaces in plain mode", () => {
    const module = bootSheetMarkdown();
    const api = module.createSheetMarkdownApi(createDeps({
      renderNarrativeBlock: (block) => block.lines.join("\n")
    }));
    const workbook = { name: "book.xlsx", sheets: [] };
    const sheet = {
      name: "Sheet1",
      index: 1,
      cells: [
        { address: "A1", row: 1, col: 1, outputValue: "Line1\nLine2", rawValue: "Raw1\nRaw2", formulaText: "", resolutionStatus: null, resolutionSource: null }
      ],
      merges: [],
      images: [],
      charts: [],
      shapes: []
    };

    const result = api.convertSheetToMarkdown(workbook, sheet, {});

    expect(result.markdown).toContain("Line1 Line2");
    expect(result.markdown).not.toContain("Line1\nLine2");
  });

  it("renders cell line breaks as <br> in github mode", () => {
    const module = bootSheetMarkdown();
    const api = module.createSheetMarkdownApi(createDeps({
      renderNarrativeBlock: (block) => block.lines.join("\n")
    }));
    const workbook = { name: "book.xlsx", sheets: [] };
    const sheet = {
      name: "Sheet1",
      index: 1,
      cells: [
        {
          address: "A1",
          row: 1,
          col: 1,
          outputValue: "Line1\nLine2",
          rawValue: "Raw1\nRaw2",
          formulaText: "",
          resolutionStatus: null,
          resolutionSource: null,
          textStyle: { bold: false, italic: false, strike: false, underline: false },
          richTextRuns: null
        }
      ],
      merges: [],
      images: [],
      charts: [],
      shapes: []
    };

    const result = api.convertSheetToMarkdown(workbook, sheet, { formattingMode: "github" });

    expect(result.markdown).toContain("Line1<br>Line2");
    expect(result.markdown).not.toContain("Line1 Line2");
  });

  it("omits shape sections when includeShapeDetails is disabled", () => {
    const module = bootSheetMarkdown();
    const api = module.createSheetMarkdownApi(createDeps({
      extractShapeBlocks: () => [{ startRow: 3, startCol: 2, endRow: 4, endCol: 3, shapeIndexes: [0] }],
      renderHierarchicalRawEntries: () => ["- kind: rect", "- text: dummy"]
    }));
    const workbook = { name: "book.xlsx", sheets: [] };
    const sheet = {
      name: "Sheet1",
      index: 1,
      cells: [],
      merges: [],
      images: [],
      charts: [],
      shapes: [{
        anchor: "B3",
        rawEntries: [{ key: "kind", value: "rect" }],
        svgFilename: null,
        svgPath: null,
        svgData: null
      }]
    };

    const enabled = api.convertSheetToMarkdown(workbook, sheet, {});
    const disabled = api.convertSheetToMarkdown(workbook, sheet, { includeShapeDetails: false });

    expect(enabled.markdown).toContain("## Shape Blocks");
    expect(enabled.markdown).toContain("## Shapes");
    expect(disabled.markdown).not.toContain("## Shape Blocks");
    expect(disabled.markdown).not.toContain("## Shapes");
  });

  it("keeps line-start markdown markers literal in narrative output", () => {
    const module = bootSheetMarkdown();
    const api = module.createSheetMarkdownApi(createDeps({
      renderNarrativeBlock: (block) => block.lines.join("\n")
    }));
    const workbook = { name: "book.xlsx", sheets: [] };
    const sheet = {
      name: "Sheet1",
      index: 1,
      cells: [
        {
          address: "A1",
          row: 1,
          col: 1,
          outputValue: "# heading",
          rawValue: "# heading",
          formulaText: "",
          resolutionStatus: null,
          resolutionSource: null,
          textStyle: { bold: false, italic: false, strike: false, underline: false },
          richTextRuns: null
        },
        {
          address: "A2",
          row: 2,
          col: 1,
          outputValue: "- item",
          rawValue: "- item",
          formulaText: "",
          resolutionStatus: null,
          resolutionSource: null,
          textStyle: { bold: false, italic: false, strike: false, underline: false },
          richTextRuns: null
        }
      ],
      merges: [],
      images: [],
      charts: [],
      shapes: []
    };

    const result = api.convertSheetToMarkdown(workbook, sheet, {});

    expect(result.markdown).toContain("\\# heading");
    expect(result.markdown).toContain("\\- item");
  });

  it("keeps ordered-list and quote markers literal in narrative output", () => {
    const module = bootSheetMarkdown();
    const api = module.createSheetMarkdownApi(createDeps({
      renderNarrativeBlock: (block) => block.lines.join("\n")
    }));
    const workbook = { name: "book.xlsx", sheets: [] };
    const sheet = {
      name: "Sheet1",
      index: 1,
      cells: [
        {
          address: "A1",
          row: 1,
          col: 1,
          outputValue: "1. item",
          rawValue: "1. item",
          formulaText: "",
          resolutionStatus: null,
          resolutionSource: null,
          textStyle: { bold: false, italic: false, strike: false, underline: false },
          richTextRuns: null
        },
        {
          address: "A2",
          row: 2,
          col: 1,
          outputValue: "> quote",
          rawValue: "> quote",
          formulaText: "",
          resolutionStatus: null,
          resolutionSource: null,
          textStyle: { bold: false, italic: false, strike: false, underline: false },
          richTextRuns: null
        }
      ],
      merges: [],
      images: [],
      charts: [],
      shapes: []
    };

    const result = api.convertSheetToMarkdown(workbook, sheet, {});

    expect(result.markdown).toContain("1\\. item");
    expect(result.markdown).toContain("&gt; quote");
  });

  it("keeps image-like markdown and code spans literal in narrative output", () => {
    const module = bootSheetMarkdown();
    const api = module.createSheetMarkdownApi(createDeps({
      renderNarrativeBlock: (block) => block.lines.join("\n")
    }));
    const workbook = { name: "book.xlsx", sheets: [] };
    const sheet = {
      name: "Sheet1",
      index: 1,
      cells: [
        {
          address: "A1",
          row: 1,
          col: 1,
          outputValue: "![alt](img.png)",
          rawValue: "![alt](img.png)",
          formulaText: "",
          resolutionStatus: null,
          resolutionSource: null,
          textStyle: { bold: false, italic: false, strike: false, underline: false },
          richTextRuns: null
        },
        {
          address: "A2",
          row: 2,
          col: 1,
          outputValue: "`code`",
          rawValue: "`code`",
          formulaText: "",
          resolutionStatus: null,
          resolutionSource: null,
          textStyle: { bold: false, italic: false, strike: false, underline: false },
          richTextRuns: null
        }
      ],
      merges: [],
      images: [],
      charts: [],
      shapes: []
    };

    const result = api.convertSheetToMarkdown(workbook, sheet, {});

    expect(result.markdown).toContain("\\!\\[alt\\]\\(img.png\\)");
    expect(result.markdown).toContain("\\`code\\`");
  });

  it("keeps additional list markers and ampersands literal in narrative output", () => {
    const module = bootSheetMarkdown();
    const api = module.createSheetMarkdownApi(createDeps({
      renderNarrativeBlock: (block) => block.lines.join("\n")
    }));
    const workbook = { name: "book.xlsx", sheets: [] };
    const sheet = {
      name: "Sheet1",
      index: 1,
      cells: [
        {
          address: "A1",
          row: 1,
          col: 1,
          outputValue: "+ plus",
          rawValue: "+ plus",
          formulaText: "",
          resolutionStatus: null,
          resolutionSource: null,
          textStyle: { bold: false, italic: false, strike: false, underline: false },
          richTextRuns: null
        },
        {
          address: "A2",
          row: 2,
          col: 1,
          outputValue: "* star",
          rawValue: "* star",
          formulaText: "",
          resolutionStatus: null,
          resolutionSource: null,
          textStyle: { bold: false, italic: false, strike: false, underline: false },
          richTextRuns: null
        },
        {
          address: "A3",
          row: 3,
          col: 1,
          outputValue: "a & b",
          rawValue: "a & b",
          formulaText: "",
          resolutionStatus: null,
          resolutionSource: null,
          textStyle: { bold: false, italic: false, strike: false, underline: false },
          richTextRuns: null
        }
      ],
      merges: [],
      images: [],
      charts: [],
      shapes: []
    };

    const result = api.convertSheetToMarkdown(workbook, sheet, {});

    expect(result.markdown).toContain("\\+ plus");
    expect(result.markdown).toContain("\\* star");
    expect(result.markdown).toContain("a &amp; b");
  });

  it("shows narrative-vs-table differences for the same markdown-like text", () => {
    const module = bootSheetMarkdown();
    const api = module.createSheetMarkdownApi(createDeps({
      detectTableCandidates: () => [{
        startRow: 2,
        startCol: 1,
        endRow: 3,
        endCol: 1,
        score: 1,
        reasonSummary: ["test"]
      }],
      matrixFromCandidate: () => [["`code` ![alt](img.png)"], ["a | b"]],
      renderNarrativeBlock: (block) => block.lines.join("\n")
    }));
    const workbook = { name: "book.xlsx", sheets: [] };
    const sheet = {
      name: "Sheet1",
      index: 1,
      cells: [
        {
          address: "A1",
          row: 1,
          col: 1,
          outputValue: "`code` ![alt](img.png)",
          rawValue: "`code` ![alt](img.png)",
          formulaText: "",
          resolutionStatus: null,
          resolutionSource: null,
          textStyle: { bold: false, italic: false, strike: false, underline: false },
          richTextRuns: null
        },
        {
          address: "A2",
          row: 2,
          col: 1,
          outputValue: "`code` ![alt](img.png)",
          rawValue: "`code` ![alt](img.png)",
          formulaText: "",
          resolutionStatus: null,
          resolutionSource: null,
          textStyle: { bold: false, italic: false, strike: false, underline: false },
          richTextRuns: null
        },
        {
          address: "A3",
          row: 3,
          col: 1,
          outputValue: "a | b",
          rawValue: "a | b",
          formulaText: "",
          resolutionStatus: null,
          resolutionSource: null,
          textStyle: { bold: false, italic: false, strike: false, underline: false },
          richTextRuns: null
        }
      ],
      merges: [],
      images: [],
      charts: [],
      shapes: []
    };

    const result = api.convertSheetToMarkdown(workbook, sheet, {});

    expect(result.markdown).toContain("\\`code\\` \\!\\[alt\\]\\(img.png\\)");
    expect(result.markdown).toContain("| `code` ![alt](img.png) |");
    expect(result.markdown).toContain("| a | b |");
  });
});
