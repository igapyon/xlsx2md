// @vitest-environment jsdom

import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";
import { loadModuleRegistry } from "./helpers/module-registry.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const zipIoCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/zip-io.js"),
  "utf8"
);
const markdownNormalizeCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/markdown-normalize.js"),
  "utf8"
);
const markdownExportCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/markdown-export.js"),
  "utf8"
);
const markdownTableEscapeCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/markdown-table-escape.js"),
  "utf8"
);

function bootMarkdownExport() {
  document.body.innerHTML = "";
  loadModuleRegistry(__dirname);
  new Function(zipIoCode)();
  new Function(markdownNormalizeCode)();
  new Function(markdownTableEscapeCode)();
  new Function(markdownExportCode)();
  return globalThis.__xlsx2mdModuleRegistry.getModule("markdownExport");
}

describe("xlsx2md markdown export", () => {
  it("normalizes line breaks into spaces", () => {
    const api = bootMarkdownExport();

    expect(api.normalizeMarkdownLineBreaks("a\r\nb\nc\rd")).toBe("a b c d");
  });

  it("renders markdown tables with escaped cell content", () => {
    const api = bootMarkdownExport();

    const markdown = api.renderMarkdownTable([
      ["Name", "Notes"],
      ["A|B", "line1\nline2"]
    ], true);

    expect(markdown).toBe(
      "| Name | Notes |\n| --- | --- |\n| A\\|B | line1 line2 |"
    );
  });

  it("delegates table cell escaping to the dedicated table-escape helper", () => {
    const api = bootMarkdownExport();

    expect(api.escapeMarkdownCell("A|\nB")).toBe("A\\| B");
  });

  it("keeps table-cell pipes and line-start markers escaped together", () => {
    const api = bootMarkdownExport();

    expect(api.escapeMarkdownCell("| a")).toBe("\\| a");
    expect(api.escapeMarkdownCell("1. item | > quote")).toBe("1. item \\| > quote");
  });

  it("normalizes html entities and pipes safely inside table cells", () => {
    const api = bootMarkdownExport();

    expect(api.escapeMarkdownCell("&lt;a&gt; &amp; b | c")).toBe("&lt;a&gt; &amp; b \\| c");
  });

  it("keeps non-pipe markdown-like text unchanged in table cells", () => {
    const api = bootMarkdownExport();

    expect(api.escapeMarkdownCell("`code` ![alt](img.png) | c")).toBe("`code` ![alt](img.png) \\| c");
  });

  it("preserves repeated spaces inside table cells while keeping pipe escaping", () => {
    const api = bootMarkdownExport();

    expect(api.escapeMarkdownCell("a   b | c")).toBe("a   b \\| c");
  });

  it("preserves leading and trailing spaces inside table cells", () => {
    const api = bootMarkdownExport();

    expect(api.escapeMarkdownCell("  a | b  ")).toBe("  a \\| b  ");
  });

  it("normalizes tabs inside table cells while keeping pipe escaping", () => {
    const api = bootMarkdownExport();

    expect(api.escapeMarkdownCell("a\tb | c")).toBe("a b \\| c");
  });

  it("creates sanitized output file names with mode suffixes", () => {
    const api = bootMarkdownExport();

    expect(api.createOutputFileName("book name.xlsx", 2, "A/B:東京", "both")).toBe(
      "book_name_002_A_B_東京_both.md"
    );
    expect(api.createOutputFileName("book name.xlsx", 2, "A/B:東京", "display", "github")).toBe(
      "book_name_002_A_B_東京_github.md"
    );
  });

  it("summarizes formula diagnostics and table scores", () => {
    const api = bootMarkdownExport();
    const summary = api.createSummaryText({
      fileName: "sample.md",
      sheetName: "Sheet1",
      markdown: "# Sheet1",
      summary: {
        outputMode: "display",
        formattingMode: "plain",
        tableDetectionMode: "balanced",
        sections: 2,
        tables: 1,
        narrativeBlocks: 1,
        merges: 0,
        images: 0,
        charts: 0,
        cells: 8,
        tableScores: [{ range: "A1-B2", score: 7, reasons: ["Has borders"] }],
        formulaDiagnostics: [
          { address: "B2", formulaText: "=A2", status: "resolved", source: "cached_value", outputValue: "1" },
          { address: "B3", formulaText: "=X1", status: "unsupported_external", source: "external_unsupported", outputValue: "" }
        ]
      }
    });

    expect(summary).toContain("Output file: sample.md");
    expect(summary).toContain("Formatting mode: plain");
    expect(summary).toContain("Table detection mode: balanced");
    expect(summary).toContain("Formula resolved: 1");
    expect(summary).toContain("Formula unsupported_external: 1");
    expect(summary).toContain("Table candidate A1-B2: score 7 / Has borders");
  });

  it("creates export entries and zip archives including markdown and assets", async () => {
    const api = bootMarkdownExport();
    const workbook = {
      name: "sample.xlsx",
      sheets: [
        {
          images: [{ path: "images/pic.png", data: new Uint8Array([1, 2, 3]) }],
          shapes: [{ svgPath: "shapes/shape_001.svg", svgData: new Uint8Array([4, 5]) }]
        }
      ]
    };
    const markdownFiles = [
      {
        fileName: "sample_001_Sheet1.md",
        sheetName: "Sheet1",
        markdown: "# Sheet1",
        summary: {
          outputMode: "display",
          formattingMode: "plain",
          tableDetectionMode: "balanced",
          sections: 1,
          tables: 0,
          narrativeBlocks: 1,
          merges: 0,
          images: 1,
          charts: 0,
          cells: 1,
          tableScores: [],
          formulaDiagnostics: []
        }
      }
    ];

    const entries = api.createExportEntries(workbook, markdownFiles);
    const archive = api.createWorkbookExportArchive(workbook, markdownFiles);
    const zipIo = globalThis.__xlsx2mdModuleRegistry.getModule("zipIo");
    const extracted = await zipIo.unzipEntries(archive.buffer.slice(archive.byteOffset, archive.byteOffset + archive.byteLength));

    expect(entries.map((entry) => entry.name).sort()).toEqual([
      "output/images/pic.png",
      "output/sample.md",
      "output/shapes/shape_001.svg"
    ]);
    expect(new TextDecoder().decode(extracted.get("output/sample.md"))).toContain("<!-- sample_001_Sheet1 -->");
    expect(extracted.get("output/images/pic.png")).toEqual(new Uint8Array([1, 2, 3]));
    expect(extracted.get("output/shapes/shape_001.svg")).toEqual(new Uint8Array([4, 5]));
  });

  it("uses formatting mode suffixes in combined export file names", () => {
    const api = bootMarkdownExport();
    const payload = api.createCombinedMarkdownExportFile(
      { name: "sample.xlsx", sheets: [{ images: [], shapes: [] }] },
      [{
        fileName: "sample_001_Sheet1_github.md",
        sheetName: "Sheet1",
        markdown: "# Sheet1",
        summary: {
          outputMode: "display",
          formattingMode: "github",
          tableDetectionMode: "balanced",
          sections: 1,
          tables: 0,
          narrativeBlocks: 1,
          merges: 0,
          images: 0,
          charts: 0,
          cells: 1,
          tableScores: [],
          formulaDiagnostics: []
        }
      }]
    );

    expect(payload.fileName).toBe("sample_github.md");
  });
});
