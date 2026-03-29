// @vitest-environment jsdom

import { Blob as NodeBlob } from "node:buffer";
import { readFileSync } from "node:fs";
import path from "node:path";
import { DecompressionStream as NodeDecompressionStream } from "node:stream/web";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

if (typeof globalThis.Blob === "undefined" || typeof globalThis.Blob.prototype?.stream !== "function") {
  globalThis.Blob = NodeBlob;
}
globalThis.DecompressionStream ??= NodeDecompressionStream;

const moduleRegistryCode = readFileSync(
  path.resolve(__dirname, "../src/js/module-registry.js"),
  "utf8"
);
const moduleRegistryAccessCode = readFileSync(
  path.resolve(__dirname, "../src/js/module-registry-access.js"),
  "utf8"
);
const runtimeEnvCode = readFileSync(
  path.resolve(__dirname, "../src/js/runtime-env.js"),
  "utf8"
);
const officeDrawingCode = readFileSync(
  path.resolve(__dirname, "../src/js/office-drawing.js"),
  "utf8"
);
const zipIoCode = readFileSync(
  path.resolve(__dirname, "../src/js/zip-io.js"),
  "utf8"
);
const borderGridCode = readFileSync(
  path.resolve(__dirname, "../src/js/border-grid.js"),
  "utf8"
);
const markdownNormalizeCode = readFileSync(
  path.resolve(__dirname, "../src/js/markdown-normalize.js"),
  "utf8"
);
const markdownEscapeCode = readFileSync(
  path.resolve(__dirname, "../src/js/markdown-escape.js"),
  "utf8"
);
const markdownTableEscapeCode = readFileSync(
  path.resolve(__dirname, "../src/js/markdown-table-escape.js"),
  "utf8"
);
const richTextParserCode = readFileSync(
  path.resolve(__dirname, "../src/js/rich-text-parser.js"),
  "utf8"
);
const richTextPlainFormatterCode = readFileSync(
  path.resolve(__dirname, "../src/js/rich-text-plain-formatter.js"),
  "utf8"
);
const richTextGithubFormatterCode = readFileSync(
  path.resolve(__dirname, "../src/js/rich-text-github-formatter.js"),
  "utf8"
);
const richTextRendererCode = readFileSync(
  path.resolve(__dirname, "../src/js/rich-text-renderer.js"),
  "utf8"
);
const narrativeStructureCode = readFileSync(
  path.resolve(__dirname, "../src/js/narrative-structure.js"),
  "utf8"
);
const tableDetectorCode = readFileSync(
  path.resolve(__dirname, "../src/js/table-detector.js"),
  "utf8"
);
const markdownExportCode = readFileSync(
  path.resolve(__dirname, "../src/js/markdown-export.js"),
  "utf8"
);
const sheetMarkdownCode = readFileSync(
  path.resolve(__dirname, "../src/js/sheet-markdown.js"),
  "utf8"
);
const stylesParserCode = readFileSync(
  path.resolve(__dirname, "../src/js/styles-parser.js"),
  "utf8"
);
const sharedStringsCode = readFileSync(
  path.resolve(__dirname, "../src/js/shared-strings.js"),
  "utf8"
);
const addressUtilsCode = readFileSync(
  path.resolve(__dirname, "../src/js/address-utils.js"),
  "utf8"
);
const relsParserCode = readFileSync(
  path.resolve(__dirname, "../src/js/rels-parser.js"),
  "utf8"
);
const worksheetTablesCode = readFileSync(
  path.resolve(__dirname, "../src/js/worksheet-tables.js"),
  "utf8"
);
const cellFormatCode = readFileSync(
  path.resolve(__dirname, "../src/js/cell-format.js"),
  "utf8"
);
const xmlUtilsCode = readFileSync(
  path.resolve(__dirname, "../src/js/xml-utils.js"),
  "utf8"
);
const sheetAssetsCode = readFileSync(
  path.resolve(__dirname, "../src/js/sheet-assets.js"),
  "utf8"
);
const worksheetParserCode = readFileSync(
  path.resolve(__dirname, "../src/js/worksheet-parser.js"),
  "utf8"
);
const workbookLoaderCode = readFileSync(
  path.resolve(__dirname, "../src/js/workbook-loader.js"),
  "utf8"
);
const formulaReferenceUtilsCode = readFileSync(
  path.resolve(__dirname, "../src/js/formula-reference-utils.js"),
  "utf8"
);
const formulaEngineCode = readFileSync(
  path.resolve(__dirname, "../src/js/formula-engine.js"),
  "utf8"
);
const formulaLegacyCode = readFileSync(
  path.resolve(__dirname, "../src/js/formula-legacy.js"),
  "utf8"
);
const formulaAstCode = readFileSync(
  path.resolve(__dirname, "../src/js/formula-ast.js"),
  "utf8"
);
const formulaResolverCode = readFileSync(
  path.resolve(__dirname, "../src/js/formula-resolver.js"),
  "utf8"
);
const coreCode = readFileSync(
  path.resolve(__dirname, "../src/js/core.js"),
  "utf8"
);

function bootCore() {
  document.body.innerHTML = "";
  new Function(moduleRegistryCode)();
  new Function(moduleRegistryAccessCode)();
  new Function(runtimeEnvCode)();
  new Function(officeDrawingCode)();
  new Function(zipIoCode)();
  new Function(borderGridCode)();
  new Function(markdownNormalizeCode)();
  new Function(markdownEscapeCode)();
  new Function(markdownTableEscapeCode)();
  new Function(richTextParserCode)();
  new Function(richTextPlainFormatterCode)();
  new Function(richTextGithubFormatterCode)();
  new Function(richTextRendererCode)();
  new Function(narrativeStructureCode)();
  new Function(tableDetectorCode)();
  new Function(markdownExportCode)();
  new Function(sheetMarkdownCode)();
  new Function(stylesParserCode)();
  new Function(sharedStringsCode)();
  new Function(addressUtilsCode)();
  new Function(relsParserCode)();
  new Function(worksheetTablesCode)();
  new Function(cellFormatCode)();
  new Function(xmlUtilsCode)();
  new Function(sheetAssetsCode)();
  new Function(worksheetParserCode)();
  new Function(workbookLoaderCode)();
  new Function(formulaReferenceUtilsCode)();
  new Function(formulaEngineCode)();
  new Function(formulaLegacyCode)();
  new Function(formulaAstCode)();
  new Function(formulaResolverCode)();
  new Function(coreCode)();
  return globalThis.__xlsx2mdModuleRegistry.getModule("xlsx2md");
}

function toArrayBuffer(bytes) {
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

async function loadFixtureMarkdown(fixtureName, formattingMode) {
  const api = bootCore();
  const fixturePath = path.resolve(__dirname, `./fixtures/rich/${fixtureName}`);
  const fixtureBytes = readFileSync(fixturePath);
  const workbook = await api.parseWorkbook(toArrayBuffer(fixtureBytes), fixtureName);
  const markdownFile = api.convertWorkbookToMarkdownFiles(workbook, {
    treatFirstRowAsHeader: true,
    trimText: true,
    removeEmptyRows: true,
    removeEmptyColumns: true,
    formattingMode
  })[0];
  return { workbook, markdownFile };
}

describe("xlsx2md rich fixtures", () => {
  it("renders rich-text-github-sample01.xlsx in github mode with inline styling and <br>", async () => {
    const { workbook, markdownFile } = await loadFixtureMarkdown("rich-text-github-sample01.xlsx", "github");

    expect(workbook.sheets).toHaveLength(1);
    expect(markdownFile.summary.formattingMode).toBe("github");
    expect(markdownFile.summary.tables).toBe(2);
    expect(markdownFile.markdown).toContain("**bold whole cell**");
    expect(markdownFile.markdown).toContain("*italic whole cell*");
    expect(markdownFile.markdown).toContain("~~strike whole cell~~");
    expect(markdownFile.markdown).toContain("<ins>underline whole cell</ins>");
    expect(markdownFile.markdown).toContain("plain **bold** *italic* strike <ins>underline</ins>");
    expect(markdownFile.markdown).toContain("***bold+italic***");
    expect(markdownFile.markdown).toContain("**<ins>bold+underline</ins>**");
    expect(markdownFile.markdown).toContain("*~~italic+strike~~*");
    expect(markdownFile.markdown).toContain("改行入り文字列で<br>**一部だけ太**字");
    expect(markdownFile.markdown).toContain("重要, <ins>下線</ins>,~~取消線~~,**強調**");
    expect(markdownFile.markdown).toContain("**12345**");
    expect(markdownFile.markdown).toContain("<ins>24690</ins>");
  });

  it("renders rich-text-github-sample01.xlsx in plain mode without inline styling or <br>", async () => {
    const { markdownFile } = await loadFixtureMarkdown("rich-text-github-sample01.xlsx", "plain");

    expect(markdownFile.summary.formattingMode).toBe("plain");
    expect(markdownFile.summary.tables).toBe(2);
    expect(markdownFile.markdown).toContain("bold whole cell");
    expect(markdownFile.markdown).toContain("italic whole cell");
    expect(markdownFile.markdown).toContain("strike whole cell");
    expect(markdownFile.markdown).toContain("underline whole cell");
    expect(markdownFile.markdown).toContain("改行入り文字列で 一部だけ太字");
    expect(markdownFile.markdown).not.toContain("<ins>underline whole cell</ins>");
    expect(markdownFile.markdown).not.toContain("**bold whole cell**");
    expect(markdownFile.markdown).not.toContain("<br>");
  });

  it("renders rich-markdown-escape-sample01.xlsx in github mode and keeps fixture-specific cases stable", async () => {
    const { workbook, markdownFile } = await loadFixtureMarkdown("rich-markdown-escape-sample01.xlsx", "github");

    expect(workbook.sheets).toHaveLength(1);
    expect(workbook.sheets[0].cells.length).toBe(36);
    expect(markdownFile.summary.formattingMode).toBe("github");
    expect(markdownFile.summary.tables).toBe(2);
    expect(markdownFile.markdown).toContain("line1 \\* x<br>**line2 \\[y\\]\\(z\\)**");
    expect(markdownFile.markdown).toContain("| Header \\\\| One | Header \\*Two\\* | Header \\[Three\\]\\(x\\) |");
    expect(markdownFile.markdown).toContain("| a**\\*b** | a\\_**b** | a\\~\\~b\\~\\~c |");
    expect(markdownFile.markdown).toContain("| \\# not **heading** | \\- not list | 1\\. ***not*** list |");
    expect(markdownFile.markdown).toContain("| a\\*b | **a\\*b** |");
    expect(markdownFile.markdown).toContain("| a\\_b | *a\\_b* |");
    expect(markdownFile.markdown).toContain("| a\\~\\~b\\~\\~c | ~~a\\~\\~b\\~\\~c~~ |");
    expect(markdownFile.markdown).toContain("| \\# not heading | <ins>\\# not heading</ins> |");
    expect(markdownFile.markdown).toContain("| &lt;tag&gt; | &lt;tag&gt; |");
    expect(markdownFile.markdown).toContain("| \\!\\[alt\\]\\(image.png\\) | \\!\\[alt\\]\\(image.png\\) |");
    expect(markdownFile.markdown).toContain("| code \\`sample\\` | code \\`sample\\` |");
    expect(markdownFile.markdown).toContain("| path\\\\to\\\\file | path\\\\to\\\\file |");
  });

  it("renders rich-markdown-escape-sample01.xlsx in plain mode as plain text without <br>", async () => {
    const { markdownFile } = await loadFixtureMarkdown("rich-markdown-escape-sample01.xlsx", "plain");

    expect(markdownFile.summary.formattingMode).toBe("plain");
    expect(markdownFile.summary.tables).toBe(2);
    expect(markdownFile.markdown).toContain("line1 \\* x line2 \\[y\\]\\(z\\)");
    expect(markdownFile.markdown).not.toContain("<br>");
    expect(markdownFile.markdown).toContain("| a\\*b | a\\*b |");
    expect(markdownFile.markdown).toContain("| a\\_b | a\\_b |");
    expect(markdownFile.markdown).toContain("| a\\~\\~b\\~\\~c | a\\~\\~b\\~\\~c |");
    expect(markdownFile.markdown).toContain("| \\# not heading | \\# not heading |");
    expect(markdownFile.markdown).toContain("| a\\*b | a\\_b | a\\~\\~b\\~\\~c |");
    expect(markdownFile.markdown).toContain("| \\# not heading | \\- not list | 1\\. not list |");
    expect(markdownFile.markdown).toContain("| \\!\\[alt\\]\\(image.png\\) | \\!\\[alt\\]\\(image.png\\) |");
    expect(markdownFile.markdown).toContain("| code \\`sample\\` | code \\`sample\\` |");
    expect(markdownFile.markdown).toContain("| path\\\\to\\\\file | path\\\\to\\\\file |");
  });

  it("renders rich-usecase-sample01.xlsx in github mode with hyperlinks, inline styling, and <br>", async () => {
    const { workbook, markdownFile } = await loadFixtureMarkdown("rich-usecase-sample01.xlsx", "github");

    expect(workbook.sheets).toHaveLength(1);
    expect(markdownFile.summary.formattingMode).toBe("github");
    expect(markdownFile.summary.tables).toBe(1);
    expect(markdownFile.markdown).toContain("| [Apple](https://www.apple.com/) | ***Apple*** の製品が<ins>購入できます</ins>。 |");
    expect(markdownFile.markdown).toContain("| [Google](https://www.google.com/) | とても<ins>有名</ins>な**検索サイト**です。 |");
    expect(markdownFile.markdown).toContain("| [Amazon](https://www.amazon.co.jp/) | **<ins>お買い物</ins>**でお世話になっています。 |");
    expect(markdownFile.markdown).toContain("| [Yodobashi](https://www.yodobashi.com/) | 実店舗とともに<br>**ネットショップ**でもお世話になっています。 |");
    expect(markdownFile.markdown).toContain("~~池袋の激戦区で、生き残るのはどの店舗か。~~<br>→トルツメ: この部分は文面から外すことを提案。");
  });

  it("renders rich-usecase-sample01.xlsx in plain mode with links preserved and styling removed", async () => {
    const { markdownFile } = await loadFixtureMarkdown("rich-usecase-sample01.xlsx", "plain");

    expect(markdownFile.summary.formattingMode).toBe("plain");
    expect(markdownFile.summary.tables).toBe(1);
    expect(markdownFile.markdown).toContain("| [Apple](https://www.apple.com/) | Apple の製品が購入できます。 |");
    expect(markdownFile.markdown).toContain("| [Google](https://www.google.com/) | とても有名な検索サイトです。 |");
    expect(markdownFile.markdown).toContain("| [Amazon](https://www.amazon.co.jp/) | お買い物でお世話になっています。 |");
    expect(markdownFile.markdown).toContain("| [Yodobashi](https://www.yodobashi.com/) | 実店舗とともに ネットショップでもお世話になっています。 |");
    expect(markdownFile.markdown).toContain("池袋の激戦区で、生き残るのはどの店舗か。 →トルツメ: この部分は文面から外すことを提案。");
    expect(markdownFile.markdown).not.toContain("<br>");
    expect(markdownFile.markdown).not.toContain("<ins>");
    expect(markdownFile.markdown).not.toContain("~~池袋の激戦区で、生き残るのはどの店舗か。~~");
  });
});
