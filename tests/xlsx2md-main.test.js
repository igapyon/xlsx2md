// @vitest-environment jsdom

import { Blob as NodeBlob } from "node:buffer";
import { readFileSync, readdirSync } from "node:fs";
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

const officeDrawingCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/office-drawing.js"),
  "utf8"
);
const moduleRegistryCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/module-registry.js"),
  "utf8"
);
const moduleRegistryAccessCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/module-registry-access.js"),
  "utf8"
);
const runtimeEnvCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/runtime-env.js"),
  "utf8"
);
const zipIoCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/zip-io.js"),
  "utf8"
);
const borderGridCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/border-grid.js"),
  "utf8"
);
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
const narrativeStructureCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/narrative-structure.js"),
  "utf8"
);
const tableDetectorCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/table-detector.js"),
  "utf8"
);
const markdownExportCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/markdown-export.js"),
  "utf8"
);
const sheetMarkdownCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/sheet-markdown.js"),
  "utf8"
);
const stylesParserCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/styles-parser.js"),
  "utf8"
);
const sharedStringsCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/shared-strings.js"),
  "utf8"
);
const addressUtilsCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/address-utils.js"),
  "utf8"
);
const relsParserCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/rels-parser.js"),
  "utf8"
);
const worksheetTablesCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/worksheet-tables.js"),
  "utf8"
);
const cellFormatCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/cell-format.js"),
  "utf8"
);
const xmlUtilsCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/xml-utils.js"),
  "utf8"
);
const sheetAssetsCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/sheet-assets.js"),
  "utf8"
);
const worksheetParserCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/worksheet-parser.js"),
  "utf8"
);
const workbookLoaderCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/workbook-loader.js"),
  "utf8"
);
const formulaReferenceUtilsCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/formula-reference-utils.js"),
  "utf8"
);
const formulaEngineCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/formula-engine.js"),
  "utf8"
);
const formulaLegacyCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/formula-legacy.js"),
  "utf8"
);
const formulaAstCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/formula-ast.js"),
  "utf8"
);
const formulaResolverCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/formula-resolver.js"),
  "utf8"
);
const coreCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/core.js"),
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

function createStoredZip(entries) {
  const encoder = new TextEncoder();
  const localChunks = [];
  const centralChunks = [];
  let offset = 0;

  for (const entry of entries) {
    const nameBytes = encoder.encode(entry.name);
    const dataBytes = typeof entry.data === "string" ? encoder.encode(entry.data) : entry.data;
    const localHeader = new Uint8Array(30 + nameBytes.length);
    const localView = new DataView(localHeader.buffer);
    localView.setUint32(0, 0x04034b50, true);
    localView.setUint16(4, 20, true);
    localView.setUint16(8, 0, true);
    localView.setUint16(10, 0, true);
    localView.setUint32(14, 0, true);
    localView.setUint32(18, dataBytes.length, true);
    localView.setUint32(22, dataBytes.length, true);
    localView.setUint16(26, nameBytes.length, true);
    localView.setUint16(28, 0, true);
    localHeader.set(nameBytes, 30);
    localChunks.push(localHeader, dataBytes);

    const centralHeader = new Uint8Array(46 + nameBytes.length);
    const centralView = new DataView(centralHeader.buffer);
    centralView.setUint32(0, 0x02014b50, true);
    centralView.setUint16(4, 20, true);
    centralView.setUint16(6, 20, true);
    centralView.setUint16(10, 0, true);
    centralView.setUint16(12, 0, true);
    centralView.setUint32(16, 0, true);
    centralView.setUint32(20, dataBytes.length, true);
    centralView.setUint32(24, dataBytes.length, true);
    centralView.setUint16(28, nameBytes.length, true);
    centralView.setUint16(30, 0, true);
    centralView.setUint16(32, 0, true);
    centralView.setUint16(34, 0, true);
    centralView.setUint16(36, 0, true);
    centralView.setUint32(38, 0, true);
    centralView.setUint32(42, offset, true);
    centralHeader.set(nameBytes, 46);
    centralChunks.push(centralHeader);

    offset += localHeader.length + dataBytes.length;
  }

  const centralStart = offset;
  const centralSize = centralChunks.reduce((sum, chunk) => sum + chunk.length, 0);
  const eocd = new Uint8Array(22);
  const eocdView = new DataView(eocd.buffer);
  eocdView.setUint32(0, 0x06054b50, true);
  eocdView.setUint16(8, entries.length, true);
  eocdView.setUint16(10, entries.length, true);
  eocdView.setUint32(12, centralSize, true);
  eocdView.setUint32(16, centralStart, true);

  const totalLength = localChunks.reduce((sum, chunk) => sum + chunk.length, 0)
    + centralSize
    + eocd.length;
  const output = new Uint8Array(totalLength);
  let cursor = 0;
  for (const chunk of localChunks) {
    output.set(chunk, cursor);
    cursor += chunk.length;
  }
  for (const chunk of centralChunks) {
    output.set(chunk, cursor);
    cursor += chunk.length;
  }
  output.set(eocd, cursor);
  return output.buffer;
}

describe("xlsx2md core", () => {
  const fixtureDir = path.resolve(__dirname, "./fixtures");

  for (const fixtureName of readdirSync(fixtureDir).filter((name) => name.toLowerCase().endsWith(".xlsx")).sort()) {
    it(`parses fixture workbook ${fixtureName} and converts it to markdown`, async () => {
      const api = bootCore();
      const fixturePath = path.resolve(fixtureDir, fixtureName);
      const fixtureBytes = readFileSync(fixturePath);
      const arrayBuffer = fixtureBytes.buffer.slice(
        fixtureBytes.byteOffset,
        fixtureBytes.byteOffset + fixtureBytes.byteLength
      );

      const workbook = await api.parseWorkbook(arrayBuffer, fixtureName);
      const files = api.convertWorkbookToMarkdownFiles(workbook, {
        treatFirstRowAsHeader: true,
        trimText: true,
        removeEmptyRows: true,
        removeEmptyColumns: true
      });

      expect(workbook.name).toBe(fixtureName);
      expect(workbook.sheets.length).toBeGreaterThan(0);
      expect(files.length).toBe(workbook.sheets.length);
      expect(files[0].markdown).toContain("## Source Information");
      expect(files[0].markdown).toContain(`Workbook: ${fixtureName}`);
    });
  }

  it("parses the basic fixture workbook with concrete baseline expectations", async () => {
    const api = bootCore();
    const fixtureName = "xlsx2md-basic-sample01.xlsx";
    const fixturePath = path.resolve(fixtureDir, fixtureName);
    const fixtureBytes = readFileSync(fixturePath);
    const arrayBuffer = fixtureBytes.buffer.slice(
      fixtureBytes.byteOffset,
      fixtureBytes.byteOffset + fixtureBytes.byteLength
    );

    const workbook = await api.parseWorkbook(arrayBuffer, fixtureName);
    const files = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true
    });
    const rawFiles = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true,
      outputMode: "raw"
    });
    const bothFiles = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true,
      outputMode: "both"
    });
    const sheet = workbook.sheets[0];
    const markdownFile = files[0];

    expect(workbook.sharedStrings.some((entry) => entry.text.includes("通常のテキスト"))).toBe(true);
    expect(workbook.sharedStrings.some((entry) => entry.text.includes("記述省略（という名前のセル結合）"))).toBe(true);
    expect(workbook.sharedStrings.length).toBe(61);

    expect(workbook.sheets).toHaveLength(1);
    expect(sheet.name).toBe("xlsx2md-basic");
    expect(sheet.maxRow).toBe(49);
    expect(sheet.maxCol).toBe(8);
    expect(sheet.cells).toHaveLength(166);
    expect(sheet.merges.map((merge) => merge.ref)).toEqual(["D28:F28", "D29:F30"]);

    expect(sheet.cells.find((cell) => cell.address === "A1")?.outputValue).toBe("通常のテキスト");
    expect(sheet.cells.find((cell) => cell.address === "B14")?.formulaText).toBe("=B13+1");
    expect(sheet.cells.find((cell) => cell.address === "B14")?.outputValue).toBe("2");
    expect(sheet.cells.find((cell) => cell.address === "B20")?.formulaText).toBe("=ROW()-19");
    expect(sheet.cells.find((cell) => cell.address === "B20")?.outputValue).toBe("1");
    expect(sheet.cells.find((cell) => cell.address === "F21")?.formulaText).toBe("=F14");
    expect(sheet.cells.find((cell) => cell.address === "F21")?.outputValue).toBe("何かの名前");
    expect(sheet.cells.find((cell) => cell.address === "E15")?.outputValue).toBe("3月13日");
    expect(sheet.cells.find((cell) => cell.address === "E16")?.outputValue).toBe("3月14日");
    expect(sheet.cells.find((cell) => cell.address === "E36")?.outputValue).toBe("1,024,768");
    expect(sheet.cells.find((cell) => cell.address === "E37")?.outputValue).toBe("¥1,024,768");
    expect(sheet.cells.find((cell) => cell.address === "E38")?.outputValue).toBe("¥ 1,024,768");
    expect(sheet.cells.find((cell) => cell.address === "E39")?.outputValue).toBe("1996/2/28");
    expect(sheet.cells.find((cell) => cell.address === "E40")?.outputValue).toBe("12:34:56");
    expect(sheet.cells.find((cell) => cell.address === "E41")?.outputValue).toBe("98.7%");
    expect(sheet.cells.find((cell) => cell.address === "E42")?.outputValue).toBe("3/4");
    expect(sheet.cells.find((cell) => cell.address === "E43")?.outputValue).toBe("1.023456E+06");
    expect(sheet.cells.find((cell) => cell.address === "E45")?.outputValue).toBe("1 0 2 3 4 5 6");
    expect(sheet.cells.find((cell) => cell.address === "E46")?.outputValue).toBe("令和8年3月17日");
    expect(sheet.cells.find((cell) => cell.address === "B37")?.formulaText).toBe("=B36+1");
    expect(sheet.cells.find((cell) => cell.address === "B40")?.formulaText).toBe("=B39+1");
    expect(sheet.cells.find((cell) => cell.address === "B41")?.formulaText).toBe("=B40+1");
    expect(sheet.cells.find((cell) => cell.address === "B46")?.formulaText).toBe("=B45+1");
    expect(sheet.cells.find((cell) => cell.address === "B46")?.outputValue).toBe("13");
    expect(sheet.cells.find((cell) => cell.address === "B46")?.resolutionStatus).toBe("resolved");
    expect(sheet.cells.find((cell) => cell.address === "D28")?.outputValue).toContain("記述省略（という名前のセル結合）");
    expect(sheet.cells.find((cell) => cell.address === "D29")?.outputValue).toContain("記述省略（2x2のセル結合）");

    expect(markdownFile.fileName).toBe("xlsx2md-basic-sample01_001_xlsx2md-basic.md");
    expect(markdownFile.summary.tables).toBe(4);
    expect(markdownFile.summary.merges).toBe(2);
    expect(markdownFile.summary.images).toBe(0);
    expect(markdownFile.summary.formulaDiagnostics).toHaveLength(27);
    expect(markdownFile.summary.formulaDiagnostics.some((diagnostic) => diagnostic.source === "cached_value")).toBe(true);
    expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual([
      "B12-F16",
      "B19-F23",
      "B26-F30",
      "B33-F46"
    ]);
    expect(markdownFile.markdown).toContain("# xlsx2md-basic");
    expect(markdownFile.markdown).toContain("Workbook: xlsx2md-basic-sample01.xlsx");
    expect(markdownFile.markdown).toContain("通常のテキスト");
    expect(markdownFile.markdown).toContain("### Table 003 (B26-F30)");
    expect(markdownFile.markdown).toContain("### Table 004 (B33-F46)");
    expect(markdownFile.markdown).toContain("| 3 | 数値 | value2 | 1,024,768 | 数値 |");
    expect(markdownFile.markdown).toContain("| 3 | 登録日 | entrydate | 3月13日 | 何かの登録日 |");
    expect(markdownFile.markdown).toContain("| 4 | 更新日 | updatedate | 3月14日 | 何かの更新日 |");
    expect(markdownFile.markdown).toContain("| 6 | 日付 | value5 | 1996/2/28 | 日付 |");
    expect(markdownFile.markdown).toContain("| 7 | 時刻 | value6 | 12:34:56 | 時刻 |");
    expect(markdownFile.markdown).toContain("| 8 | パーセンテージ | value7 | 98.7% | パーセンテージ |");
    expect(markdownFile.markdown).toContain("| 9 | 分数 | value8 | 3/4 | 分数 |");
    expect(markdownFile.markdown).toContain("| 10 | 指数 | value9 | 1.023456E+06 | 指数 |");
    expect(markdownFile.markdown).toContain("| 12 | その他 | value11 | 1 0 2 3 4 5 6 | その他 |");
    expect(markdownFile.markdown).toContain("| 13 | ユーザー定義 | value12 | 令和8年3月17日 | ユーザー定義 |");
    expect(markdownFile.markdown).toContain("[MERGED←]");
    expect(markdownFile.markdown).toContain("[MERGED↑]");
  });

  it("parses the display-format fixture workbook with concrete display expectations", async () => {
    const api = bootCore();
    const fixtureName = "display-format-sample01.xlsx";
    const fixturePath = path.resolve(fixtureDir, "display", fixtureName);
    const fixtureBytes = readFileSync(fixturePath);
    const arrayBuffer = fixtureBytes.buffer.slice(
      fixtureBytes.byteOffset,
      fixtureBytes.byteOffset + fixtureBytes.byteLength
    );

    const workbook = await api.parseWorkbook(arrayBuffer, fixtureName);
    const files = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true
    });
    const sheet = workbook.sheets[0];
    const markdownFile = files[0];

    expect(workbook.sheets).toHaveLength(1);
    expect(sheet.name).toBe("display-format");
    expect(sheet.maxRow).toBe(13);
    expect(sheet.maxCol).toBe(5);
    expect(sheet.cells).toHaveLength(65);

    expect(sheet.cells.find((cell) => cell.address === "D3")?.outputValue).toBe("1,024,768");
    expect(sheet.cells.find((cell) => cell.address === "D4")?.outputValue).toBe("¥1,024,768");
    expect(sheet.cells.find((cell) => cell.address === "D5")?.outputValue).toBe("¥ 1,024,768");
    expect(sheet.cells.find((cell) => cell.address === "D6")?.outputValue).toBe("1996/2/28");
    expect(sheet.cells.find((cell) => cell.address === "D7")?.outputValue).toBe("12:34:56");
    expect(sheet.cells.find((cell) => cell.address === "D8")?.outputValue).toBe("98.7%");
    expect(sheet.cells.find((cell) => cell.address === "D9")?.outputValue).toBe("3/4");
    expect(sheet.cells.find((cell) => cell.address === "D10")?.outputValue).toBe("1.023456E+06");
    expect(sheet.cells.find((cell) => cell.address === "D11")?.outputValue).toBe("1023456");
    expect(sheet.cells.find((cell) => cell.address === "D12")?.outputValue).toBe("1023456");
    expect(sheet.cells.find((cell) => cell.address === "D13")?.outputValue).toBe("令和8年3月17日");

    expect(markdownFile.fileName).toBe("display-format-sample01_001_display-format.md");
    expect(markdownFile.summary.tables).toBe(1);
    expect(markdownFile.summary.tableScores).toHaveLength(1);
    expect(markdownFile.summary.tableScores[0].range).toBe("A1-E13");
    expect(markdownFile.markdown).toContain("# display-format");
    expect(markdownFile.markdown).toContain("Workbook: display-format-sample01.xlsx");
    expect(markdownFile.markdown).toContain("| 2 | 数値 | value2 | 1,024,768 | 数値 |");
    expect(markdownFile.markdown).toContain("| 3 | 通貨 | value3 | ¥1,024,768 | 通貨 |");
    expect(markdownFile.markdown).toContain("| 4 | 会計 | value4 | ¥ 1,024,768 | 会計 |");
    expect(markdownFile.markdown).toContain("| 5 | 日付 | value5 | 1996/2/28 | 日付 |");
    expect(markdownFile.markdown).toContain("| 8 | 分数 | value8 | 3/4 | 分数 |");
    expect(markdownFile.markdown).toContain("| 12 | 和暦 | value12 | 令和8年3月17日 | 和暦 |");
  });

  it("extracts chart metadata blocks from drawing charts", async () => {
    const api = bootCore();
    const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <sheets>
          <sheet name="ChartSheet" sheetId="1" r:id="rId1"/>
        </sheets>
      </workbook>`;
    const workbookRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
      </Relationships>`;
    const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <sheetData>
          <row r="1"><c r="A1" t="inlineStr"><is><t>Chart Sample</t></is></c></row>
        </sheetData>
        <drawing r:id="rId1"/>
      </worksheet>`;
    const sheetRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
      </Relationships>`;
    const drawingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <xdr:twoCellAnchor>
          <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>17</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
          <xdr:to><xdr:col>7</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>34</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
          <xdr:graphicFrame>
            <xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
            <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                <c:chart r:id="rId1"/>
              </a:graphicData>
            </a:graphic>
          </xdr:graphicFrame>
          <xdr:clientData/>
        </xdr:twoCellAnchor>
      </xdr:wsDr>`;
    const drawingRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
      </Relationships>`;
    const chartXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <c:chart>
          <c:title><c:tx><c:rich><a:p><a:r><a:t>Sample Sales Chart</a:t></a:r></a:p></c:rich></c:tx></c:title>
          <c:plotArea>
            <c:barChart>
              <c:ser>
                <c:tx><c:v>Series A</c:v></c:tx>
                <c:cat><c:strRef><c:f>ChartSheet!$B$3:$B$6</c:f></c:strRef></c:cat>
                <c:val><c:numRef><c:f>ChartSheet!$E$3:$E$6</c:f></c:numRef></c:val>
              </c:ser>
              <c:ser>
                <c:tx><c:v>Series B</c:v></c:tx>
                <c:cat><c:strRef><c:f>ChartSheet!$B$3:$B$6</c:f></c:strRef></c:cat>
                <c:val><c:numRef><c:f>ChartSheet!$D$3:$D$6</c:f></c:numRef></c:val>
              </c:ser>
            </c:barChart>
          </c:plotArea>
        </c:chart>
      </c:chartSpace>`;
    const arrayBuffer = createStoredZip([
      { name: "[Content_Types].xml", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
        </Types>` },
      { name: "_rels/.rels", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
        </Relationships>` },
      { name: "xl/workbook.xml", data: workbookXml },
      { name: "xl/_rels/workbook.xml.rels", data: workbookRels },
      { name: "xl/worksheets/sheet1.xml", data: sheetXml },
      { name: "xl/worksheets/_rels/sheet1.xml.rels", data: sheetRels },
      { name: "xl/drawings/drawing1.xml", data: drawingXml },
      { name: "xl/drawings/_rels/drawing1.xml.rels", data: drawingRels },
      { name: "xl/charts/chart1.xml", data: chartXml }
    ]);

    const workbook = await api.parseWorkbook(arrayBuffer, "chart-sample.xlsx");
    const markdownFile = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true
    })[0];

    expect(workbook.sheets[0].charts).toHaveLength(1);
    expect(workbook.sheets[0].charts[0].anchor).toBe("B18");
    expect(workbook.sheets[0].charts[0].title).toBe("Sample Sales Chart");
    expect(workbook.sheets[0].charts[0].chartType).toContain("Bar Chart");
    expect(workbook.sheets[0].charts[0].series).toHaveLength(2);
    expect(workbook.sheets[0].charts[0].series[0].axis).toBe("primary");
    expect(markdownFile.summary.charts).toBe(1);
    expect(markdownFile.markdown).toContain("## Charts");
    expect(markdownFile.markdown).toContain("### Chart 001 (B18)");
    expect(markdownFile.markdown).toContain("- Title: Sample Sales Chart");
    expect(markdownFile.markdown).toContain("- Type: Bar Chart");
    expect(markdownFile.markdown).toContain("  - Series A");
    expect(markdownFile.markdown).toContain("    - categories: ChartSheet!$B$3:$B$6");
    expect(markdownFile.markdown).toContain("    - values: ChartSheet!$E$3:$E$6");
  });

  it("parses the merge-pattern fixture workbook with concrete merge expectations", async () => {
    const api = bootCore();
    const fixtureName = "merge-pattern-sample01.xlsx";
    const fixturePath = path.resolve(fixtureDir, "merge", fixtureName);
    const fixtureBytes = readFileSync(fixturePath);
    const arrayBuffer = fixtureBytes.buffer.slice(
      fixtureBytes.byteOffset,
      fixtureBytes.byteOffset + fixtureBytes.byteLength
    );

    const workbook = await api.parseWorkbook(arrayBuffer, fixtureName);
    const files = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true
    });
    const sheet = workbook.sheets[0];
    const markdownFile = files[0];

    expect(workbook.sheets).toHaveLength(1);
    expect(sheet.name).toBe("merge");
    expect(sheet.maxRow).toBe(19);
    expect(sheet.maxCol).toBe(5);
    expect(sheet.cells).toHaveLength(68);
    expect(sheet.merges.map((merge) => merge.ref)).toEqual([
      "B15:C16",
      "B17:C18",
      "D15:E16",
      "D17:E18",
      "B2:C2",
      "B3:D3",
      "B4:E4",
      "D2:E2",
      "B8:B9",
      "B10:B11",
      "C8:C10",
      "D8:D11"
    ]);

    expect(sheet.cells.find((cell) => cell.address === "B2")?.outputValue).toBe("横結合");
    expect(sheet.cells.find((cell) => cell.address === "C2")?.outputValue).toBe("");
    expect(sheet.cells.find((cell) => cell.address === "D2")?.outputValue).toBe("横結合");
    expect(sheet.cells.find((cell) => cell.address === "E2")?.outputValue).toBe("");
    expect(sheet.cells.find((cell) => cell.address === "A9")?.formulaText).toBe("=A8+1");
    expect(sheet.cells.find((cell) => cell.address === "A9")?.outputValue).toBe("2");
    expect(sheet.cells.find((cell) => cell.address === "B9")?.outputValue).toBe("");
    expect(sheet.cells.find((cell) => cell.address === "C9")?.outputValue).toBe("");
    expect(sheet.cells.find((cell) => cell.address === "D11")?.outputValue).toBe("");
    expect(sheet.cells.find((cell) => cell.address === "B15")?.outputValue).toBe("2x2結合");
    expect(sheet.cells.find((cell) => cell.address === "C15")?.outputValue).toBe("");
    expect(sheet.cells.find((cell) => cell.address === "E16")?.outputValue).toBe("");
    expect(sheet.cells.find((cell) => cell.address === "C18")?.outputValue).toBe("");

    expect(markdownFile.fileName).toBe("merge-pattern-sample01_001_merge.md");
    expect(markdownFile.summary.tables).toBe(3);
    expect(markdownFile.summary.merges).toBe(12);
    expect(markdownFile.summary.images).toBe(0);
    expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual([
      "A1-E4",
      "A7-D11",
      "A14-E18"
    ]);
    expect(markdownFile.markdown).toContain("# merge");
    expect(markdownFile.markdown).toContain("Workbook: merge-pattern-sample01.xlsx");
    expect(markdownFile.markdown).toContain("### Table 001 (A1-E4)");
    expect(markdownFile.markdown).toContain("### Table 002 (A7-D11)");
    expect(markdownFile.markdown).toContain("### Table 003 (A14-E18)");
    expect(markdownFile.markdown).toContain("| 1 | 横結合 | [MERGED←] | 横結合 | [MERGED←] |");
    expect(markdownFile.markdown).toContain("| 2 | [MERGED↑] | [MERGED↑] | [MERGED↑] |");
    expect(markdownFile.markdown).toContain("| 1 | 2x2結合 | [MERGED←] | 2x2結合 | [MERGED←] |");
    expect(markdownFile.markdown).toContain("※横結合のサンプルです");
    expect(markdownFile.markdown).toContain("※縦結合のサンプルです");
    expect(markdownFile.markdown).toContain("※2x2結合のサンプルです");
    expect(markdownFile.markdown).toContain("[MERGED←]");
    expect(markdownFile.markdown).toContain("[MERGED↑]");
  });

  it("parses the merge-multiline fixture workbook with concrete multiline-merge expectations", async () => {
    const api = bootCore();
    const fixtureName = "merge-multiline-sample01.xlsx";
    const fixturePath = path.resolve(fixtureDir, "merge", fixtureName);
    const fixtureBytes = readFileSync(fixturePath);
    const arrayBuffer = fixtureBytes.buffer.slice(
      fixtureBytes.byteOffset,
      fixtureBytes.byteOffset + fixtureBytes.byteLength
    );

    const workbook = await api.parseWorkbook(arrayBuffer, fixtureName);
    const files = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true
    });
    const sheet = workbook.sheets[0];
    const markdownFile = files[0];

    expect(workbook.sheets).toHaveLength(1);
    expect(sheet.name).toBe("merge-multiline");
    expect(sheet.maxRow).toBe(6);
    expect(sheet.maxCol).toBe(3);
    expect(sheet.cells).toHaveLength(11);
    expect(sheet.merges.map((merge) => merge.ref)).toEqual(["B3:C4"]);

    expect(sheet.cells.find((cell) => cell.address === "B3")?.outputValue).toBe("1行目\n2行目");
    expect(sheet.cells.find((cell) => cell.address === "C3")?.outputValue).toBe("");
    expect(sheet.cells.find((cell) => cell.address === "B4")?.outputValue).toBe("");
    expect(sheet.cells.find((cell) => cell.address === "A6")?.outputValue).toBe("※結合セル内の改行確認用");

    expect(markdownFile.fileName).toBe("merge-multiline-sample01_001_merge-multiline.md");
    expect(markdownFile.summary.tables).toBe(1);
    expect(markdownFile.summary.merges).toBe(1);
    expect(markdownFile.summary.images).toBe(0);
    expect(markdownFile.summary.formulaDiagnostics).toHaveLength(0);
    expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual(["A1-C4"]);
    expect(markdownFile.markdown).toContain("# merge-multiline");
    expect(markdownFile.markdown).toContain("Workbook: merge-multiline-sample01.xlsx");
    expect(markdownFile.markdown).toContain("### Table 001 (A1-C4)");
    expect(markdownFile.markdown).toContain("| 1 | 1行目 2行目 | [MERGED←] |");
    expect(markdownFile.markdown).toContain("| 2 | [MERGED↑] | [MERGED↑] |");
    expect(markdownFile.markdown).toContain("※結合セル内の改行確認用");
  });

  it("parses the formula-basic fixture workbook with concrete formula expectations", async () => {
    const api = bootCore();
    const fixtureName = "formula-basic-sample01.xlsx";
    const fixturePath = path.resolve(fixtureDir, "formula", fixtureName);
    const fixtureBytes = readFileSync(fixturePath);
    const arrayBuffer = fixtureBytes.buffer.slice(
      fixtureBytes.byteOffset,
      fixtureBytes.byteOffset + fixtureBytes.byteLength
    );

    const workbook = await api.parseWorkbook(arrayBuffer, fixtureName);
    const files = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true
    });
    const sheet = workbook.sheets[0];
    const markdownFile = files[0];

    expect(workbook.sheets).toHaveLength(1);
    expect(sheet.name).toBe("formula");
    expect(sheet.maxRow).toBe(13);
    expect(sheet.maxCol).toBe(2);
    expect(sheet.cells).toHaveLength(23);

    expect(sheet.cells.find((cell) => cell.address === "B5")?.formulaText).toBe("=B3");
    expect(sheet.cells.find((cell) => cell.address === "B5")?.outputValue).toBe("10");
    expect(sheet.cells.find((cell) => cell.address === "B5")?.resolutionStatus).toBe("resolved");
    expect(sheet.cells.find((cell) => cell.address === "B6")?.formulaText).toBe("=B3+B4");
    expect(sheet.cells.find((cell) => cell.address === "B6")?.outputValue).toBe("15");
    expect(sheet.cells.find((cell) => cell.address === "B7")?.formulaText).toBe("=IF(B3>B4,\"OK\",\"NG\")");
    expect(sheet.cells.find((cell) => cell.address === "B7")?.outputValue).toBe("OK");
    expect(sheet.cells.find((cell) => cell.address === "B8")?.formulaText).toBe("=SUM(B3:B4)");
    expect(sheet.cells.find((cell) => cell.address === "B8")?.outputValue).toBe("15");
    expect(sheet.cells.find((cell) => cell.address === "B9")?.formulaText).toBe("=COUNTIF(B3:B4,\">7\")");
    expect(sheet.cells.find((cell) => cell.address === "B9")?.outputValue).toBe("1");
    expect(sheet.cells.find((cell) => cell.address === "B10")?.formulaText).toBe("=TEXT(B3,\"0000\")");
    expect(sheet.cells.find((cell) => cell.address === "B10")?.outputValue).toBe("0010");
    expect(sheet.cells.find((cell) => cell.address === "B11")?.formulaText).toBe("=DATE(2024,3,17)");
    expect(sheet.cells.find((cell) => cell.address === "B11")?.outputValue).toBe("2024/3/17");
    expect(sheet.cells.find((cell) => cell.address === "B12")?.formulaText).toBe("=VALUE(\"1,234.5\")");
    expect(sheet.cells.find((cell) => cell.address === "B12")?.outputValue).toBe("1234.5");
    expect(sheet.cells.find((cell) => cell.address === "B13")?.formulaText).toBe("=VALUE(\"2024/03/17\")");
    expect(sheet.cells.find((cell) => cell.address === "B13")?.outputValue).toBe("45368");

    expect(markdownFile.fileName).toBe("formula-basic-sample01_001_formula.md");
    expect(markdownFile.summary.tables).toBe(1);
    expect(markdownFile.summary.merges).toBe(0);
    expect(markdownFile.summary.images).toBe(0);
    expect(markdownFile.summary.formulaDiagnostics).toHaveLength(9);
    expect(markdownFile.summary.formulaDiagnostics.every((diagnostic) => diagnostic.source === "cached_value")).toBe(true);
    expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual(["A3-B13"]);
    expect(markdownFile.markdown).toContain("# formula");
    expect(markdownFile.markdown).toContain("Workbook: formula-basic-sample01.xlsx");
    expect(markdownFile.markdown).toContain("| ref | 10 |");
    expect(markdownFile.markdown).toContain("| arith | 15 |");
    expect(markdownFile.markdown).toContain("| if | OK |");
    expect(markdownFile.markdown).toContain("| sum | 15 |");
    expect(markdownFile.markdown).toContain("| countif | 1 |");
    expect(markdownFile.markdown).toContain("| text | 0010 |");
    expect(markdownFile.markdown).toContain("| date | 2024/3/17 |");
    expect(markdownFile.markdown).toContain("| value\\_num | 1234.5 |");
    expect(markdownFile.markdown).toContain("| value\\_date | 45368 |");
  });

  it("parses the formula-spill fixture workbook with concrete spill-like expectations", async () => {
    const api = bootCore();
    const fixtureName = "formula-spill-sample01.xlsx";
    const fixturePath = path.resolve(fixtureDir, "formula", fixtureName);
    const fixtureBytes = readFileSync(fixturePath);
    const arrayBuffer = fixtureBytes.buffer.slice(
      fixtureBytes.byteOffset,
      fixtureBytes.byteOffset + fixtureBytes.byteLength
    );

    const workbook = await api.parseWorkbook(arrayBuffer, fixtureName);
    const files = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true
    });
    const sheet = workbook.sheets[0];
    const markdownFile = files[0];

    expect(workbook.sheets).toHaveLength(1);
    expect(sheet.name).toBe("spill-sample");
    expect(sheet.maxRow).toBe(6);
    expect(sheet.maxCol).toBe(5);
    expect(sheet.cells).toHaveLength(11);

    expect(sheet.cells.find((cell) => cell.address === "A4")?.outputValue).toBe("1");
    expect(sheet.cells.find((cell) => cell.address === "A6")?.outputValue).toBe("3");
    expect(sheet.cells.find((cell) => cell.address === "C4")?.formulaText).toBe("=_xlfn.SEQUENCE(3)");
    expect(sheet.cells.find((cell) => cell.address === "C4")?.outputValue).toBe("1");
    expect(sheet.cells.find((cell) => cell.address === "C4")?.resolutionStatus).toBe("resolved");
    expect(sheet.cells.find((cell) => cell.address === "C5")?.outputValue).toBe("2");
    expect(sheet.cells.find((cell) => cell.address === "C6")?.outputValue).toBe("3");
    expect(sheet.cells.find((cell) => cell.address === "E4")?.formulaText).toBe("=SUM(_xlfn.ANCHORARRAY(C4))");
    expect(sheet.cells.find((cell) => cell.address === "E4")?.outputValue).toBe("6");
    expect(sheet.cells.find((cell) => cell.address === "E4")?.resolutionStatus).toBe("resolved");

    expect(markdownFile.fileName).toBe("formula-spill-sample01_001_spill-sample.md");
    expect(markdownFile.summary.tables).toBe(0);
    expect(markdownFile.summary.merges).toBe(0);
    expect(markdownFile.summary.images).toBe(0);
    expect(markdownFile.summary.formulaDiagnostics).toHaveLength(2);
    expect(markdownFile.summary.formulaDiagnostics.every((diagnostic) => diagnostic.source === "cached_value")).toBe(true);
    expect(markdownFile.summary.tableScores).toHaveLength(0);
    expect(markdownFile.markdown).toContain("# spill-sample");
    expect(markdownFile.markdown).toContain("Workbook: formula-spill-sample01.xlsx");
    expect(markdownFile.markdown).toContain("spill サンプル");
    expect(markdownFile.markdown).toContain("src1 spill\\_ref spill\\_sum");
    expect(markdownFile.markdown).toContain("1 1 6");
    expect(markdownFile.markdown).toContain("2 2");
    expect(markdownFile.markdown).toContain("3 3");
  });

  it("parses the formula-crosssheet fixture workbook with concrete cross-sheet expectations", async () => {
    const api = bootCore();
    const fixtureName = "formula-crosssheet-sample01.xlsx";
    const fixturePath = path.resolve(fixtureDir, "formula", fixtureName);
    const fixtureBytes = readFileSync(fixturePath);
    const arrayBuffer = fixtureBytes.buffer.slice(
      fixtureBytes.byteOffset,
      fixtureBytes.byteOffset + fixtureBytes.byteLength
    );

    const workbook = await api.parseWorkbook(arrayBuffer, fixtureName);
    const files = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true
    });
    const sheet1 = workbook.sheets[0];
    const sheet2 = workbook.sheets[1];
    const sheet3 = workbook.sheets[2];
    const sheet1File = files[0];
    const sheet2File = files[1];
    const sheet3File = files[2];

    expect(workbook.sheets).toHaveLength(3);

    expect(sheet1.name).toBe("Sheet1");
    expect(sheet1.maxRow).toBe(13);
    expect(sheet1.maxCol).toBe(2);
    expect(sheet1.cells).toHaveLength(23);
    expect(sheet1.cells.find((cell) => cell.address === "A1")?.outputValue).toBe("複数シート参照サンプル");
    expect(sheet1.cells.find((cell) => cell.address === "B3")?.formulaText).toBe("=Sheet2!B3");
    expect(sheet1.cells.find((cell) => cell.address === "B3")?.outputValue).toBe("CrossValue");
    expect(sheet1.cells.find((cell) => cell.address === "B3")?.resolutionStatus).toBe("resolved");
    expect(sheet1.cells.find((cell) => cell.address === "B4")?.formulaText).toBe("=日本語シート!C4");
    expect(sheet1.cells.find((cell) => cell.address === "B4")?.outputValue).toBe("日本語参照値");
    expect(sheet1.cells.find((cell) => cell.address === "B4")?.resolutionStatus).toBe("resolved");
    expect(sheet1.cells.find((cell) => cell.address === "B5")?.formulaText).toBe("=SUM(Sheet2!A1:B2)");
    expect(sheet1.cells.find((cell) => cell.address === "B5")?.outputValue).toBe("10");
    expect(sheet1.cells.find((cell) => cell.address === "B5")?.resolutionStatus).toBe("resolved");

    expect(sheet2.name).toBe("Sheet2");
    expect(sheet2.maxRow).toBe(3);
    expect(sheet2.maxCol).toBe(2);
    expect(sheet2.cells).toHaveLength(5);
    expect(sheet2.cells.find((cell) => cell.address === "B3")?.outputValue).toBe("CrossValue");

    expect(sheet3.name).toBe("日本語シート");
    expect(sheet3.maxRow).toBe(4);
    expect(sheet3.maxCol).toBe(3);
    expect(sheet3.cells).toHaveLength(1);
    expect(sheet3.cells.find((cell) => cell.address === "C4")?.outputValue).toBe("日本語参照値");

    expect(sheet1File.fileName).toBe("formula-crosssheet-sample01_001_Sheet1.md");
    expect(sheet1File.summary.tables).toBe(1);
    expect(sheet1File.summary.tableScores.map((detail) => detail.range)).toEqual(["A3-B5"]);
    expect(sheet1File.summary.formulaDiagnostics).toHaveLength(3);
    expect(sheet1File.summary.formulaDiagnostics.every((diagnostic) => diagnostic.source === "cached_value")).toBe(true);
    expect(sheet1File.markdown).toContain("# Sheet1");
    expect(sheet1File.markdown).toContain("Workbook: formula-crosssheet-sample01.xlsx");
    expect(sheet1File.markdown).toContain("| sheet2\\_ref | CrossValue |");
    expect(sheet1File.markdown).toContain("| jp\\_sheet\\_ref | 日本語参照値 |");
    expect(sheet1File.markdown).toContain("| sum\\_range | 10 |");

    expect(sheet2File.fileName).toBe("formula-crosssheet-sample01_002_Sheet2.md");
    expect(sheet2File.summary.tables).toBe(1);
    expect(sheet2File.summary.tableScores.map((detail) => detail.range)).toEqual(["A1-B3"]);
    expect(sheet2File.markdown).toContain("| 1 | 2 |");
    expect(sheet2File.markdown).toContain("| 3 | 4 |");
    expect(sheet2File.markdown).toContain("|  | CrossValue |");

    expect(sheet3File.fileName).toBe("formula-crosssheet-sample01_003_日本語シート.md");
    expect(sheet3File.summary.tables).toBe(0);
    expect(sheet3File.summary.tableScores).toHaveLength(0);
    expect(sheet3File.markdown).toContain("# 日本語シート");
    expect(sheet3File.markdown).toContain("日本語参照値");
  });

  it("parses the image-basic fixture workbook with concrete image expectations", async () => {
    const api = bootCore();
    const fixtureName = "image-basic-sample01.xlsx";
    const fixturePath = path.resolve(fixtureDir, "image", fixtureName);
    const fixtureBytes = readFileSync(fixturePath);
    const arrayBuffer = fixtureBytes.buffer.slice(
      fixtureBytes.byteOffset,
      fixtureBytes.byteOffset + fixtureBytes.byteLength
    );

    const workbook = await api.parseWorkbook(arrayBuffer, fixtureName);
    const files = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true
    });
    const sheet = workbook.sheets[0];
    const markdownFile = files[0];

    expect(workbook.sheets).toHaveLength(1);
    expect(sheet.name).toBe("image");
    expect(sheet.maxRow).toBe(13);
    expect(sheet.maxCol).toBe(6);
    expect(sheet.cells).toHaveLength(39);
    expect(sheet.images).toHaveLength(2);
    expect(sheet.images.map((image) => ({
      filename: image.filename,
      path: image.path,
      anchor: image.anchor,
      mediaPath: image.mediaPath
    }))).toEqual([
      {
        filename: "image_001.png",
        path: "assets/image/image_001.png",
        anchor: "C8",
        mediaPath: "xl/media/image1.png"
      },
      {
        filename: "image_002.png",
        path: "assets/image/image_002.png",
        anchor: "F8",
        mediaPath: "xl/media/image2.png"
      }
    ]);
    expect(sheet.cells.find((cell) => cell.address === "A1")?.outputValue).toBe("画像抽出サンプル");
    expect(sheet.cells.find((cell) => cell.address === "D4")?.outputValue).toBe("123,456");
    expect(sheet.cells.find((cell) => cell.address === "E4")?.outputValue).toBe("1,234.56");
    expect(sheet.cells.find((cell) => cell.address === "F4")?.outputValue).toBe("値1");
    expect(sheet.cells.find((cell) => cell.address === "D6")?.outputValue).toBe("345,678");
    expect(sheet.cells.find((cell) => cell.address === "E6")?.outputValue).toBe("3,456.89");

    expect(markdownFile.fileName).toBe("image-basic-sample01_001_image.md");
    expect(markdownFile.summary.tables).toBe(1);
    expect(markdownFile.summary.merges).toBe(0);
    expect(markdownFile.summary.images).toBe(2);
    expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual(["B3-F6"]);
    expect(markdownFile.markdown).toContain("# image");
    expect(markdownFile.markdown).toContain("Workbook: image-basic-sample01.xlsx");
    expect(markdownFile.markdown).toContain("| 項目1 | 項目2 | 項目3 | 項目4 | 項目5 |");
    expect(markdownFile.markdown).toContain("| B4 | C4 | 123,456 | 1,234.56 | 値1 |");
    expect(markdownFile.markdown).toContain("| B6 | C6 | 345,678 | 3,456.89 | 値3 |");
    expect(markdownFile.markdown).toContain("## Images");
    expect(markdownFile.markdown).toContain("### Image 001 (C8)");
    expect(markdownFile.markdown).toContain("- File: assets/image/image_001.png");
    expect(markdownFile.markdown).toContain("![image_001.png](assets/image/image_001.png)");
    expect(markdownFile.markdown).toContain("### Image 002 (F8)");
    expect(markdownFile.markdown).toContain("- File: assets/image/image_002.png");
    expect(markdownFile.markdown).toContain("![image_002.png](assets/image/image_002.png)");
  });

  it("parses the image-basic-sample02 fixture workbook with concrete image and chart expectations", async () => {
    const api = bootCore();
    const fixtureName = "image-basic-sample02.xlsx";
    const fixturePath = path.resolve(fixtureDir, "image", fixtureName);
    const fixtureBytes = readFileSync(fixturePath);
    const arrayBuffer = fixtureBytes.buffer.slice(
      fixtureBytes.byteOffset,
      fixtureBytes.byteOffset + fixtureBytes.byteLength
    );

    const workbook = await api.parseWorkbook(arrayBuffer, fixtureName);
    const files = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true
    });
    const sheet = workbook.sheets[0];
    const markdownFile = files[0];

    expect(workbook.sheets).toHaveLength(1);
    expect(sheet.name).toBe("image");
    expect(sheet.maxRow).toBe(6);
    expect(sheet.maxCol).toBe(4);
    expect(sheet.cells).toHaveLength(13);
    expect(sheet.images).toHaveLength(1);
    expect(sheet.charts).toHaveLength(1);
    expect(sheet.images[0]).toMatchObject({
      filename: "image_001.png",
      path: "assets/image/image_001.png",
      anchor: "H3",
      mediaPath: "xl/media/image1.png"
    });
    expect(sheet.charts[0]).toEqual({
      sheetName: "image",
      anchor: "B9",
      chartPath: "xl/charts/chart1.xml",
      title: "このグラフのタイトル",
      chartType: "Line Chart",
      series: [
        {
          name: "値A",
          categoriesRef: "image!$B$4:$B$6",
          valuesRef: "image!$C$4:$C$6",
          axis: "primary"
        },
        {
          name: "値B",
          categoriesRef: "image!$B$4:$B$6",
          valuesRef: "image!$D$4:$D$6",
          axis: "primary"
        }
      ]
    });

    expect(markdownFile.fileName).toBe("image-basic-sample02_001_image.md");
    expect(markdownFile.summary.tables).toBe(1);
    expect(markdownFile.summary.images).toBe(1);
    expect(markdownFile.summary.charts).toBe(1);
    expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual(["B3-D6"]);
    expect(markdownFile.markdown).toContain("# image");
    expect(markdownFile.markdown).toContain("Workbook: image-basic-sample02.xlsx");
    expect(markdownFile.markdown).toContain("| 項目 | 値A | 値B |");
    expect(markdownFile.markdown).toContain("| 2024年 | 13,568 | 9,072 |");
    expect(markdownFile.markdown).toContain("## Charts");
    expect(markdownFile.markdown).toContain("### Chart 001 (B9)");
    expect(markdownFile.markdown).toContain("- Title: このグラフのタイトル");
    expect(markdownFile.markdown).toContain("- Type: Line Chart");
    expect(markdownFile.markdown).toContain("    - values: image!$D$4:$D$6");
    expect(markdownFile.markdown).toContain("## Images");
    expect(markdownFile.markdown).toContain("### Image 001 (H3)");
    expect(markdownFile.markdown).toContain("- File: assets/image/image_001.png");
  });

  it("parses the chart-basic fixture workbook with concrete chart expectations", async () => {
    const api = bootCore();
    const fixtureName = "chart-basic-sample01.xlsx";
    const fixturePath = path.resolve(fixtureDir, "chart", fixtureName);
    const fixtureBytes = readFileSync(fixturePath);
    const arrayBuffer = fixtureBytes.buffer.slice(
      fixtureBytes.byteOffset,
      fixtureBytes.byteOffset + fixtureBytes.byteLength
    );

    const workbook = await api.parseWorkbook(arrayBuffer, fixtureName);
    const files = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true
    });
    const sheet = workbook.sheets[0];
    const markdownFile = files[0];

    expect(workbook.sheets).toHaveLength(1);
    expect(sheet.name).toBe("chart-basic");
    expect(sheet.maxRow).toBe(7);
    expect(sheet.maxCol).toBe(4);
    expect(sheet.cells).toHaveLength(16);
    expect(sheet.images).toHaveLength(0);
    expect(sheet.charts).toHaveLength(1);
    expect(sheet.charts[0]).toEqual({
      sheetName: "chart-basic",
      anchor: "B10",
      chartPath: "xl/charts/chart1.xml",
      title: "棒グラフのグラフ",
      chartType: "Bar Chart",
      series: [
        {
          name: "値A",
          categoriesRef: "'chart-basic'!$B$4:$B$7",
          valuesRef: "'chart-basic'!$C$4:$C$7",
          axis: "primary"
        },
        {
          name: "値B",
          categoriesRef: "'chart-basic'!$B$4:$B$7",
          valuesRef: "'chart-basic'!$D$4:$D$7",
          axis: "primary"
        }
      ]
    });

    expect(markdownFile.fileName).toBe("chart-basic-sample01_001_chart-basic.md");
    expect(markdownFile.summary.tables).toBe(1);
    expect(markdownFile.summary.images).toBe(0);
    expect(markdownFile.summary.charts).toBe(1);
    expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual(["B3-D7"]);
    expect(markdownFile.markdown).toContain("# chart-basic");
    expect(markdownFile.markdown).toContain("Workbook: chart-basic-sample01.xlsx");
    expect(markdownFile.markdown).toContain("グラフ基本サンプル");
    expect(markdownFile.markdown).toContain("### Table 001 (B3-D7)");
    expect(markdownFile.markdown).toContain("| 項目 | 値A | 値B |");
    expect(markdownFile.markdown).toContain("| 2027年 | 28,053 | 32,012 |");
    expect(markdownFile.markdown).toContain("## Charts");
    expect(markdownFile.markdown).toContain("### Chart 001 (B10)");
    expect(markdownFile.markdown).toContain("- Title: 棒グラフのグラフ");
    expect(markdownFile.markdown).toContain("- Type: Bar Chart");
    expect(markdownFile.markdown).toContain("    - categories: 'chart-basic'!$B$4:$B$7");
    expect(markdownFile.markdown).toContain("    - values: 'chart-basic'!$D$4:$D$7");
  });

  it("parses the chart-mixed fixture workbook with concrete mixed-chart expectations", async () => {
    const api = bootCore();
    const fixtureName = "chart-mixed-sample01.xlsx";
    const fixturePath = path.resolve(fixtureDir, "chart", fixtureName);
    const fixtureBytes = readFileSync(fixturePath);
    const arrayBuffer = fixtureBytes.buffer.slice(
      fixtureBytes.byteOffset,
      fixtureBytes.byteOffset + fixtureBytes.byteLength
    );

    const workbook = await api.parseWorkbook(arrayBuffer, fixtureName);
    const files = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true
    });
    const sheet = workbook.sheets[0];
    const markdownFile = files[0];

    expect(workbook.sheets).toHaveLength(1);
    expect(sheet.name).toBe("chart-mixed");
    expect(sheet.maxRow).toBe(8);
    expect(sheet.maxCol).toBe(5);
    expect(sheet.cells).toHaveLength(25);
    expect(sheet.images).toHaveLength(0);
    expect(sheet.charts).toHaveLength(1);
    expect(sheet.charts[0]).toEqual({
      sheetName: "chart-mixed",
      anchor: "B10",
      chartPath: "xl/charts/chart1.xml",
      title: "棒と折れ線",
      chartType: "Bar Chart + Line Chart (Combined)",
      series: [
        {
          name: "売上",
          categoriesRef: "'chart-mixed'!$B$4:$B$8",
          valuesRef: "'chart-mixed'!$C$4:$C$8",
          axis: "primary"
        },
        {
          name: "割引額",
          categoriesRef: "'chart-mixed'!$B$4:$B$8",
          valuesRef: "'chart-mixed'!$D$4:$D$8",
          axis: "primary"
        },
        {
          name: "利益率",
          categoriesRef: "'chart-mixed'!$B$4:$B$8",
          valuesRef: "'chart-mixed'!$E$4:$E$8",
          axis: "secondary"
        }
      ]
    });

    expect(markdownFile.fileName).toBe("chart-mixed-sample01_001_chart-mixed.md");
    expect(markdownFile.summary.tables).toBe(1);
    expect(markdownFile.summary.images).toBe(0);
    expect(markdownFile.summary.charts).toBe(1);
    expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual(["B3-E8"]);
    expect(markdownFile.markdown).toContain("# chart-mixed");
    expect(markdownFile.markdown).toContain("Workbook: chart-mixed-sample01.xlsx");
    expect(markdownFile.markdown).toContain("| 項目 | 売上 | 割引額 | 利益率 |");
    expect(markdownFile.markdown).toContain("| 2028年 | 31,027 | 2,500 | 18% |");
    expect(markdownFile.markdown).toContain("## Charts");
    expect(markdownFile.markdown).toContain("### Chart 001 (B10)");
    expect(markdownFile.markdown).toContain("- Title: 棒と折れ線");
    expect(markdownFile.markdown).toContain("- Type: Bar Chart + Line Chart (Combined)");
    expect(markdownFile.markdown).toContain("  - 利益率");
    expect(markdownFile.markdown).toContain("    - Axis: secondary");
    expect(markdownFile.markdown).toContain("    - values: 'chart-mixed'!$E$4:$E$8");
  });

  it("parses the shape-basic fixture workbook without misclassifying drawing shapes", async () => {
    const api = bootCore();
    const fixtureName = "shape-basic-sample01.xlsx";
    const fixturePath = path.resolve(fixtureDir, "shape", fixtureName);
    const fixtureBytes = readFileSync(fixturePath);
    const arrayBuffer = fixtureBytes.buffer.slice(
      fixtureBytes.byteOffset,
      fixtureBytes.byteOffset + fixtureBytes.byteLength
    );

    const workbook = await api.parseWorkbook(arrayBuffer, fixtureName);
    const files = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true
    });
    const sheet = workbook.sheets[0];
    const markdownFile = files[0];

    expect(workbook.sheets).toHaveLength(1);
    expect(sheet.name).toBe("shape-basic");
    expect(sheet.maxRow).toBe(6);
    expect(sheet.maxCol).toBe(5);
    expect(sheet.cells).toHaveLength(17);
    expect(sheet.images).toHaveLength(0);
    expect(sheet.charts).toHaveLength(0);
    expect(sheet.shapes).toHaveLength(3);
    expect(sheet.shapes[0]).toMatchObject({
      sheetName: "shape-basic",
      anchor: "H3",
      name: "テキスト ボックス 1",
      kind: "Text Box",
      text: "テキストボックスの例",
      widthEmu: 1980029,
      heightEmu: 392608,
      elementName: "xdr:sp",
      svgFilename: "shape_001.svg",
      svgPath: "assets/shape-basic/shape_001.svg"
    });
    expect(sheet.shapes[0].rawEntries).toEqual(expect.arrayContaining([
      { key: "xdr:oneCellAnchor/xdr:from/xdr:col#text", value: "7" },
      { key: "xdr:oneCellAnchor/xdr:from/xdr:row#text", value: "2" },
      { key: "xdr:oneCellAnchor/xdr:sp/xdr:nvSpPr/xdr:cNvPr@name", value: "テキスト ボックス 1" },
      { key: "xdr:oneCellAnchor/xdr:sp/xdr:nvSpPr/xdr:cNvSpPr@txBox", value: "1" },
      { key: "xdr:oneCellAnchor/xdr:sp/xdr:txBody/a:p/a:r/a:t#text", value: "テキストボックスの例" }
    ]));
    expect(sheet.shapes[1]).toMatchObject({
      sheetName: "shape-basic",
      anchor: "H8",
      name: "直線矢印コネクタ 3",
      kind: "Straight Arrow Connector",
      text: "",
      widthEmu: 1308100,
      heightEmu: 0,
      elementName: "xdr:cxnSp",
      svgFilename: "shape_002.svg",
      svgPath: "assets/shape-basic/shape_002.svg"
    });
    expect(sheet.shapes[1].rawEntries).toEqual(expect.arrayContaining([
      { key: "xdr:twoCellAnchor/xdr:cxnSp/xdr:nvCxnSpPr/xdr:cNvPr@name", value: "直線矢印コネクタ 3" },
      { key: "xdr:twoCellAnchor/xdr:cxnSp/xdr:spPr/a:prstGeom@prst", value: "straightConnector1" }
    ]));
    expect(sheet.shapes[2]).toMatchObject({
      sheetName: "shape-basic",
      anchor: "K3",
      name: "正方形/長方形 4",
      kind: "Rectangle",
      text: "",
      widthEmu: 1511300,
      heightEmu: 1155700,
      elementName: "xdr:sp",
      svgFilename: "shape_003.svg",
      svgPath: "assets/shape-basic/shape_003.svg"
    });
    expect(sheet.shapes[2].rawEntries).toEqual(expect.arrayContaining([
      { key: "xdr:twoCellAnchor/xdr:sp/xdr:nvSpPr/xdr:cNvPr@name", value: "正方形/長方形 4" },
      { key: "xdr:twoCellAnchor/xdr:sp/xdr:spPr/a:prstGeom@prst", value: "rect" },
      { key: "xdr:twoCellAnchor/xdr:sp/xdr:spPr/a:xfrm/a:ext@cx", value: "1511300" },
      { key: "xdr:twoCellAnchor/xdr:sp/xdr:spPr/a:xfrm/a:ext@cy", value: "1155700" }
    ]));

    expect(markdownFile.fileName).toBe("shape-basic-sample01_001_shape-basic.md");
    expect(markdownFile.summary.tables).toBe(1);
    expect(markdownFile.summary.images).toBe(0);
    expect(markdownFile.summary.charts).toBe(0);
    expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual(["B3-E6"]);
    expect(markdownFile.markdown).toContain("# shape-basic");
    expect(markdownFile.markdown).toContain("Workbook: shape-basic-sample01.xlsx");
    expect(markdownFile.markdown).toContain("図形サンプル");
    expect(markdownFile.markdown).toContain("### Table 001 (B3-E6)");
    expect(markdownFile.markdown).toContain("| 項目 | 値A | 値B | 値C |");
    expect(markdownFile.markdown).toContain("| 2026年 | 25,051 | 32,012 | 14,850 |");
    expect(markdownFile.markdown).toContain("## Shape Blocks");
    expect(markdownFile.markdown).toContain("### Shape Block 001 (");
    expect(markdownFile.markdown).toContain("- Shapes: Shape 001, Shape 002, Shape 003");
    expect(markdownFile.markdown).toContain("## Shapes");
    expect(markdownFile.markdown).toContain("### Shape 001 (H3)");
    expect(markdownFile.markdown).toContain("- `xdr:oneCellAnchor`");
    expect(markdownFile.markdown).toContain("    - `xdr:from`");
        expect(markdownFile.markdown).toContain("        - `xdr:col#text`: `7`");
    expect(markdownFile.markdown).toContain("    - `xdr:sp`");
    expect(markdownFile.markdown).toContain("- `xdr:cNvPr@name`: `テキスト ボックス 1`");
    expect(markdownFile.markdown).toContain("- `a:t#text`: `テキストボックスの例`");
    expect(markdownFile.markdown).toContain("- SVG: assets/shape-basic/shape_001.svg");
    expect(markdownFile.markdown).toContain("![shape_001.svg](assets/shape-basic/shape_001.svg)");
    expect(markdownFile.markdown).toContain("### Shape 002 (H8)");
    expect(markdownFile.markdown).toContain("- `xdr:twoCellAnchor`");
    expect(markdownFile.markdown).toContain("    - `xdr:cxnSp`");
    expect(markdownFile.markdown).toContain("- `a:prstGeom@prst`: `straightConnector1`");
    expect(markdownFile.markdown).toContain("- `a:ext@cx`: `1308100`");
    expect(markdownFile.markdown).toContain("### Shape 003 (K3)");
    expect(markdownFile.markdown).toContain("- `xdr:cNvPr@name`: `正方形/長方形 4`");
    expect(markdownFile.markdown).toContain("![shape_003.svg](assets/shape-basic/shape_003.svg)");
    expect(markdownFile.markdown).not.toContain("## Images");
    expect(markdownFile.markdown).not.toContain("## Charts");
  });

  it("exports shape SVG assets into the markdown+assets archive", async () => {
    const api = bootCore();
    const fixtureName = "shape-basic-sample01.xlsx";
    const fixturePath = path.resolve(fixtureDir, "shape", fixtureName);
    const fixtureBytes = readFileSync(fixturePath);
    const arrayBuffer = fixtureBytes.buffer.slice(
      fixtureBytes.byteOffset,
      fixtureBytes.byteOffset + fixtureBytes.byteLength
    );

    const workbook = await api.parseWorkbook(arrayBuffer, fixtureName);
    const files = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true
    });
    const archive = api.createWorkbookExportArchive(workbook, files);
    const extracted = await api.unzipEntries(
      archive.buffer.slice(archive.byteOffset, archive.byteOffset + archive.byteLength)
    );

    expect(extracted.has("output/shape-basic-sample01.md")).toBe(true);
    expect(extracted.has("output/assets/shape-basic/shape_001.svg")).toBe(true);
    expect(extracted.has("output/assets/shape-basic/shape_002.svg")).toBe(true);
    expect(extracted.has("output/assets/shape-basic/shape_003.svg")).toBe(true);

    const markdownText = new TextDecoder().decode(extracted.get("output/shape-basic-sample01.md"));
    expect(markdownText).toContain("## Shapes");
    expect(markdownText).toContain("![shape_001.svg](assets/shape-basic/shape_001.svg)");

    const shape1Svg = new TextDecoder().decode(extracted.get("output/assets/shape-basic/shape_001.svg"));
    const shape2Svg = new TextDecoder().decode(extracted.get("output/assets/shape-basic/shape_002.svg"));
    const shape3Svg = new TextDecoder().decode(extracted.get("output/assets/shape-basic/shape_003.svg"));
    expect(shape1Svg).toContain("<svg");
    expect(shape1Svg).toContain("<text");
    expect(shape2Svg).toContain("<svg");
    expect(shape2Svg).toContain("<line");
    expect(shape3Svg).toContain("<svg");
    expect(shape3Svg).toContain("<rect");
  });

  it("parses the shape-flowchart fixture workbook with flowchart raw metadata and connector SVG export", async () => {
    const api = bootCore();
    const fixtureName = "shape-flowchart-sample01.xlsx";
    const fixturePath = path.resolve(fixtureDir, "shape", fixtureName);
    const fixtureBytes = readFileSync(fixturePath);
    const arrayBuffer = fixtureBytes.buffer.slice(
      fixtureBytes.byteOffset,
      fixtureBytes.byteOffset + fixtureBytes.byteLength
    );

    const workbook = await api.parseWorkbook(arrayBuffer, fixtureName);
    const files = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true
    });
    const sheet = workbook.sheets[0];
    const markdownFile = files[0];

    expect(workbook.sheets).toHaveLength(1);
    expect(sheet.name).toBe("shape-flowchart");
    expect(sheet.maxRow).toBe(6);
    expect(sheet.maxCol).toBe(7);
    expect(sheet.cells).toHaveLength(18);
    expect(sheet.images).toHaveLength(0);
    expect(sheet.charts).toHaveLength(0);
    expect(sheet.shapes).toHaveLength(7);

    expect(sheet.shapes[0]).toMatchObject({
      sheetName: "shape-flowchart",
      anchor: "H3",
      name: "フローチャート: 端子 2",
      kind: "Shape (flowChartTerminator)",
      text: "開始",
      widthEmu: 1689100,
      heightEmu: 584200,
      svgPath: null
    });
    expect(sheet.shapes[0].rawEntries).toEqual(expect.arrayContaining([
      { key: "xdr:twoCellAnchor/xdr:sp/xdr:nvSpPr/xdr:cNvPr@name", value: "フローチャート: 端子 2" },
      { key: "xdr:twoCellAnchor/xdr:sp/xdr:spPr/a:prstGeom@prst", value: "flowChartTerminator" },
      { key: "xdr:twoCellAnchor/xdr:sp/xdr:txBody/a:p/a:r/a:t#text", value: "開始" }
    ]));
    expect(sheet.shapes[1]).toMatchObject({
      anchor: "K3",
      name: "フローチャート: 処理 5",
      kind: "Shape (flowChartProcess)",
      text: "処理",
      svgPath: null
    });
    expect(sheet.shapes[1].rawEntries).toEqual(expect.arrayContaining([
      { key: "xdr:twoCellAnchor/xdr:sp/xdr:spPr/a:prstGeom@prst", value: "flowChartProcess" },
      { key: "xdr:twoCellAnchor/xdr:sp/xdr:txBody/a:p/a:r/a:t#text", value: "処理" }
    ]));
    expect(sheet.shapes[2]).toMatchObject({
      anchor: "N3",
      name: "フローチャート: 判断 6",
      kind: "Shape (flowChartDecision)",
      text: "条件判断",
      svgPath: null
    });
    expect(sheet.shapes[2].rawEntries).toEqual(expect.arrayContaining([
      { key: "xdr:twoCellAnchor/xdr:sp/xdr:spPr/a:prstGeom@prst", value: "flowChartDecision" },
      { key: "xdr:twoCellAnchor/xdr:sp/xdr:txBody/a:p/a:r/a:t#text", value: "条件判断" }
    ]));
    expect(sheet.shapes[3]).toMatchObject({
      anchor: "Q3",
      name: "フローチャート: データ 7",
      kind: "Shape (flowChartInputOutput)",
      text: "データ",
      svgPath: null
    });
    expect(sheet.shapes[3].rawEntries).toEqual(expect.arrayContaining([
      { key: "xdr:twoCellAnchor/xdr:sp/xdr:spPr/a:prstGeom@prst", value: "flowChartInputOutput" },
      { key: "xdr:twoCellAnchor/xdr:sp/xdr:txBody/a:p/a:r/a:t#text", value: "データ" }
    ]));
    expect(sheet.shapes[4]).toMatchObject({
      anchor: "I4",
      name: "直線矢印コネクタ 9",
      kind: "Straight Arrow Connector",
      text: "",
      svgFilename: "shape_005.svg",
      svgPath: "assets/shape-flowchart/shape_005.svg"
    });
    expect(sheet.shapes[4].rawEntries).toEqual(expect.arrayContaining([
      { key: "xdr:twoCellAnchor/xdr:cxnSp/xdr:spPr/a:prstGeom@prst", value: "straightConnector1" },
      { key: "xdr:twoCellAnchor/xdr:cxnSp/xdr:nvCxnSpPr/xdr:cNvCxnSpPr/a:endCxn@id", value: "6" }
    ]));

    expect(markdownFile.fileName).toBe("shape-flowchart-sample01_001_shape-flowchart.md");
    expect(markdownFile.summary.tables).toBe(1);
    expect(markdownFile.summary.images).toBe(0);
    expect(markdownFile.summary.charts).toBe(0);
    expect(markdownFile.summary.formulaDiagnostics).toHaveLength(0);
    expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual(["B3-E6"]);
    expect(markdownFile.markdown).toContain("# shape-flowchart");
    expect(markdownFile.markdown).toContain("Workbook: shape-flowchart-sample01.xlsx");
    expect(markdownFile.markdown).toContain("フローチャート図形サンプル");
    expect(markdownFile.markdown).toContain("### Table 001 (B3-E6)");
    expect(markdownFile.markdown).toContain("## Shape Blocks");
    expect(markdownFile.markdown).toContain("### Shape Block 001 (H3-S7)");
    expect(markdownFile.markdown).toContain("- Shapes: Shape 001, Shape 002, Shape 003, Shape 004, Shape 005, Shape 006, Shape 007");
    expect(markdownFile.markdown).toContain("## Shapes");
    expect(markdownFile.markdown).toContain("### Shape 001 (H3)");
    expect(markdownFile.markdown).toContain("- `a:prstGeom@prst`: `flowChartTerminator`");
    expect(markdownFile.markdown).toContain("- `a:t#text`: `開始`");
    expect(markdownFile.markdown).toContain("### Shape 003 (N3)");
    expect(markdownFile.markdown).toContain("- `a:prstGeom@prst`: `flowChartDecision`");
    expect(markdownFile.markdown).toContain("- `a:t#text`: `条件判断`");
    expect(markdownFile.markdown).toContain("### Shape 005 (I4)");
    expect(markdownFile.markdown).toContain("- `a:prstGeom@prst`: `straightConnector1`");
    expect(markdownFile.markdown).toContain("- `a:endCxn@id`: `6`");
    expect(markdownFile.markdown).toContain("![shape_005.svg](assets/shape-flowchart/shape_005.svg)");
    expect(markdownFile.markdown).toContain("![shape_006.svg](assets/shape-flowchart/shape_006.svg)");
    expect(markdownFile.markdown).toContain("![shape_007.svg](assets/shape-flowchart/shape_007.svg)");
    expect(markdownFile.markdown).not.toContain("## Images");
    expect(markdownFile.markdown).not.toContain("## Charts");
  });

  it("exports flowchart connector SVG assets into the markdown+assets archive", async () => {
    const api = bootCore();
    const fixtureName = "shape-flowchart-sample01.xlsx";
    const fixturePath = path.resolve(fixtureDir, "shape", fixtureName);
    const fixtureBytes = readFileSync(fixturePath);
    const arrayBuffer = fixtureBytes.buffer.slice(
      fixtureBytes.byteOffset,
      fixtureBytes.byteOffset + fixtureBytes.byteLength
    );

    const workbook = await api.parseWorkbook(arrayBuffer, fixtureName);
    const files = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true
    });
    const archive = api.createWorkbookExportArchive(workbook, files);
    const extracted = await api.unzipEntries(
      archive.buffer.slice(archive.byteOffset, archive.byteOffset + archive.byteLength)
    );

    expect(extracted.has("output/shape-flowchart-sample01.md")).toBe(true);
    expect(extracted.has("output/assets/shape-flowchart/shape_005.svg")).toBe(true);
    expect(extracted.has("output/assets/shape-flowchart/shape_006.svg")).toBe(true);
    expect(extracted.has("output/assets/shape-flowchart/shape_007.svg")).toBe(true);
    expect(extracted.has("output/assets/shape-flowchart/shape_001.svg")).toBe(false);

    const markdownText = new TextDecoder().decode(extracted.get("output/shape-flowchart-sample01.md"));
    expect(markdownText).toContain("## Shapes");
    expect(markdownText).toContain("![shape_005.svg](assets/shape-flowchart/shape_005.svg)");

    const shape5Svg = new TextDecoder().decode(extracted.get("output/assets/shape-flowchart/shape_005.svg"));
    const shape6Svg = new TextDecoder().decode(extracted.get("output/assets/shape-flowchart/shape_006.svg"));
    const shape7Svg = new TextDecoder().decode(extracted.get("output/assets/shape-flowchart/shape_007.svg"));
    expect(shape5Svg).toContain("<svg");
    expect(shape5Svg).toContain("<line");
    expect(shape6Svg).toContain("<svg");
    expect(shape6Svg).toContain("<line");
    expect(shape7Svg).toContain("<svg");
    expect(shape7Svg).toContain("<line");
  });

  it("parses the shape-block-arrow fixture workbook with block-arrow raw metadata", async () => {
    const api = bootCore();
    const fixtureName = "shape-block-arrow-sample01.xlsx";
    const fixturePath = path.resolve(fixtureDir, "shape", fixtureName);
    const fixtureBytes = readFileSync(fixturePath);
    const arrayBuffer = fixtureBytes.buffer.slice(
      fixtureBytes.byteOffset,
      fixtureBytes.byteOffset + fixtureBytes.byteLength
    );

    const workbook = await api.parseWorkbook(arrayBuffer, fixtureName);
    const files = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true
    });
    const sheet = workbook.sheets[0];
    const markdownFile = files[0];

    expect(workbook.sheets).toHaveLength(1);
    expect(sheet.name).toBe("shape-block-arrow");
    expect(sheet.maxRow).toBe(6);
    expect(sheet.maxCol).toBe(7);
    expect(sheet.cells).toHaveLength(18);
    expect(sheet.images).toHaveLength(0);
    expect(sheet.charts).toHaveLength(0);
    expect(sheet.shapes).toHaveLength(5);

    expect(sheet.shapes[0]).toMatchObject({
      sheetName: "shape-block-arrow",
      anchor: "H3",
      name: "右矢印 22",
      kind: "Shape (rightArrow)",
      text: "右矢印",
      widthEmu: 2108200,
      heightEmu: 1066800,
      svgPath: null
    });
    expect(sheet.shapes[0].rawEntries).toEqual(expect.arrayContaining([
      { key: "xdr:twoCellAnchor/xdr:sp/xdr:nvSpPr/xdr:cNvPr@name", value: "右矢印 22" },
      { key: "xdr:twoCellAnchor/xdr:sp/xdr:spPr/a:prstGeom@prst", value: "rightArrow" },
      { key: "xdr:twoCellAnchor/xdr:sp/xdr:txBody/a:p/a:r/a:t#text", value: "右矢印" }
    ]));
    expect(sheet.shapes[1]).toMatchObject({
      anchor: "K3",
      name: "左右矢印 24",
      kind: "Shape (leftRightArrow)",
      text: "",
      svgPath: null
    });
    expect(sheet.shapes[1].rawEntries).toEqual(expect.arrayContaining([
      { key: "xdr:twoCellAnchor/xdr:sp/xdr:spPr/a:prstGeom@prst", value: "leftRightArrow" }
    ]));
    expect(sheet.shapes[2]).toMatchObject({
      anchor: "N3",
      name: "上矢印 25",
      kind: "Shape (upArrow)",
      text: "",
      svgPath: null
    });
    expect(sheet.shapes[2].rawEntries).toEqual(expect.arrayContaining([
      { key: "xdr:twoCellAnchor/xdr:sp/xdr:spPr/a:prstGeom@prst", value: "upArrow" }
    ]));
    expect(sheet.shapes[3]).toMatchObject({
      anchor: "Q3",
      name: "U ターン矢印 26",
      kind: "Shape (uturnArrow)",
      text: "",
      svgPath: null
    });
    expect(sheet.shapes[3].rawEntries).toEqual(expect.arrayContaining([
      { key: "xdr:twoCellAnchor/xdr:sp/xdr:spPr/a:prstGeom@prst", value: "uturnArrow" }
    ]));
    expect(sheet.shapes[4]).toMatchObject({
      anchor: "H8",
      name: "四方向矢印 27",
      kind: "Shape (quadArrow)",
      text: "",
      svgPath: null
    });
    expect(sheet.shapes[4].rawEntries).toEqual(expect.arrayContaining([
      { key: "xdr:twoCellAnchor/xdr:sp/xdr:spPr/a:prstGeom@prst", value: "quadArrow" }
    ]));

    expect(markdownFile.fileName).toBe("shape-block-arrow-sample01_001_shape-block-arrow.md");
    expect(markdownFile.summary.tables).toBe(1);
    expect(markdownFile.summary.images).toBe(0);
    expect(markdownFile.summary.charts).toBe(0);
    expect(markdownFile.summary.formulaDiagnostics).toHaveLength(0);
    expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual(["B3-E6"]);
    expect(markdownFile.markdown).toContain("# shape-block-arrow");
    expect(markdownFile.markdown).toContain("Workbook: shape-block-arrow-sample01.xlsx");
    expect(markdownFile.markdown).toContain("ブロック矢印サンプル");
    expect(markdownFile.markdown).toContain("## Shape Blocks");
    expect(markdownFile.markdown).toContain("### Shape Block 001 (H3-S14)");
    expect(markdownFile.markdown).toContain("- Shapes: Shape 001, Shape 002, Shape 003, Shape 004, Shape 005");
    expect(markdownFile.markdown).toContain("## Shapes");
    expect(markdownFile.markdown).toContain("### Shape 001 (H3)");
    expect(markdownFile.markdown).toContain("- `a:prstGeom@prst`: `rightArrow`");
    expect(markdownFile.markdown).toContain("- `a:t#text`: `右矢印`");
    expect(markdownFile.markdown).toContain("### Shape 002 (K3)");
    expect(markdownFile.markdown).toContain("- `a:prstGeom@prst`: `leftRightArrow`");
    expect(markdownFile.markdown).toContain("### Shape 003 (N3)");
    expect(markdownFile.markdown).toContain("- `a:prstGeom@prst`: `upArrow`");
    expect(markdownFile.markdown).toContain("### Shape 004 (Q3)");
    expect(markdownFile.markdown).toContain("- `a:prstGeom@prst`: `uturnArrow`");
    expect(markdownFile.markdown).toContain("### Shape 005 (H8)");
    expect(markdownFile.markdown).toContain("- `a:prstGeom@prst`: `quadArrow`");
    expect(markdownFile.markdown).not.toContain("![shape_001.svg]");
    expect(markdownFile.markdown).not.toContain("## Images");
    expect(markdownFile.markdown).not.toContain("## Charts");
  });

  it("exports only markdown for the block-arrow fixture archive when no SVG assets are generated", async () => {
    const api = bootCore();
    const fixtureName = "shape-block-arrow-sample01.xlsx";
    const fixturePath = path.resolve(fixtureDir, "shape", fixtureName);
    const fixtureBytes = readFileSync(fixturePath);
    const arrayBuffer = fixtureBytes.buffer.slice(
      fixtureBytes.byteOffset,
      fixtureBytes.byteOffset + fixtureBytes.byteLength
    );

    const workbook = await api.parseWorkbook(arrayBuffer, fixtureName);
    const files = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true
    });
    const archive = api.createWorkbookExportArchive(workbook, files);
    const extracted = await api.unzipEntries(
      archive.buffer.slice(archive.byteOffset, archive.byteOffset + archive.byteLength)
    );

    expect(extracted.has("output/shape-block-arrow-sample01.md")).toBe(true);
    expect(Array.from(extracted.keys())).toEqual(["output/shape-block-arrow-sample01.md"]);

    const markdownText = new TextDecoder().decode(extracted.get("output/shape-block-arrow-sample01.md"));
    expect(markdownText).toContain("## Shapes");
    expect(markdownText).toContain("- `a:prstGeom@prst`: `uturnArrow`");
    expect(markdownText).not.toContain("![shape_001.svg]");
  });

  it("parses the shape-callout fixture workbook with callout raw metadata and text extraction", async () => {
    const api = bootCore();
    const fixtureName = "shape-callout-sample01.xlsx";
    const fixturePath = path.resolve(fixtureDir, "shape", fixtureName);
    const fixtureBytes = readFileSync(fixturePath);
    const arrayBuffer = fixtureBytes.buffer.slice(
      fixtureBytes.byteOffset,
      fixtureBytes.byteOffset + fixtureBytes.byteLength
    );

    const workbook = await api.parseWorkbook(arrayBuffer, fixtureName);
    const files = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true
    });
    const sheet = workbook.sheets[0];
    const markdownFile = files[0];

    expect(workbook.sheets).toHaveLength(1);
    expect(sheet.name).toBe("shape-callout");
    expect(sheet.maxRow).toBe(6);
    expect(sheet.maxCol).toBe(7);
    expect(sheet.cells).toHaveLength(18);
    expect(sheet.images).toHaveLength(0);
    expect(sheet.charts).toHaveLength(0);
    expect(sheet.shapes).toHaveLength(4);

    expect(sheet.shapes[0]).toMatchObject({
      sheetName: "shape-callout",
      anchor: "H3",
      name: "角丸四角形吹き出し 2",
      kind: "Shape (wedgeRoundRectCallout)",
      text: "角四角",
      widthEmu: 2374900,
      heightEmu: 901700,
      svgPath: null
    });
    expect(sheet.shapes[0].rawEntries).toEqual(expect.arrayContaining([
      { key: "xdr:twoCellAnchor/xdr:sp/xdr:nvSpPr/xdr:cNvPr@name", value: "角丸四角形吹き出し 2" },
      { key: "xdr:twoCellAnchor/xdr:sp/xdr:spPr/a:prstGeom@prst", value: "wedgeRoundRectCallout" },
      { key: "xdr:twoCellAnchor/xdr:sp/xdr:txBody/a:p/a:r/a:t#text", value: "角四角" }
    ]));
    expect(sheet.shapes[1]).toMatchObject({
      anchor: "K3",
      name: "円形吹き出し 3",
      kind: "Shape (wedgeEllipseCallout)",
      text: "楕円",
      svgPath: null
    });
    expect(sheet.shapes[1].rawEntries).toEqual(expect.arrayContaining([
      { key: "xdr:twoCellAnchor/xdr:sp/xdr:spPr/a:prstGeom@prst", value: "wedgeEllipseCallout" },
      { key: "xdr:twoCellAnchor/xdr:sp/xdr:txBody/a:p/a:r/a:t#text", value: "楕円" }
    ]));
    expect(sheet.shapes[2]).toMatchObject({
      anchor: "N3",
      name: "雲形吹き出し 4",
      kind: "Shape (cloudCallout)",
      text: "雲",
      svgPath: null
    });
    expect(sheet.shapes[2].rawEntries).toEqual(expect.arrayContaining([
      { key: "xdr:twoCellAnchor/xdr:sp/xdr:spPr/a:prstGeom@prst", value: "cloudCallout" },
      { key: "xdr:twoCellAnchor/xdr:sp/xdr:txBody/a:p/a:r/a:t#text", value: "雲" }
    ]));
    expect(sheet.shapes[3]).toMatchObject({
      anchor: "H8",
      name: "強調線吹き出し 1 (枠付き) 1",
      kind: "Shape (accentBorderCallout1)",
      text: "注釈",
      svgPath: null
    });
    expect(sheet.shapes[3].rawEntries).toEqual(expect.arrayContaining([
      { key: "xdr:twoCellAnchor/xdr:sp/xdr:spPr/a:prstGeom@prst", value: "accentBorderCallout1" },
      { key: "xdr:twoCellAnchor/xdr:sp/xdr:txBody/a:p/a:r/a:t#text", value: "注釈" }
    ]));

    expect(markdownFile.fileName).toBe("shape-callout-sample01_001_shape-callout.md");
    expect(markdownFile.summary.tables).toBe(1);
    expect(markdownFile.summary.images).toBe(0);
    expect(markdownFile.summary.charts).toBe(0);
    expect(markdownFile.summary.formulaDiagnostics).toHaveLength(0);
    expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual(["B3-E6"]);
    expect(markdownFile.markdown).toContain("# shape-callout");
    expect(markdownFile.markdown).toContain("Workbook: shape-callout-sample01.xlsx");
    expect(markdownFile.markdown).toContain("吹き出しサンプル");
    expect(markdownFile.markdown).toContain("## Shape Blocks");
    expect(markdownFile.markdown).toContain("### Shape Block 001 (H3-P12)");
    expect(markdownFile.markdown).toContain("- Shapes: Shape 001, Shape 002, Shape 003, Shape 004");
    expect(markdownFile.markdown).toContain("## Shapes");
    expect(markdownFile.markdown).toContain("### Shape 001 (H3)");
    expect(markdownFile.markdown).toContain("- `a:prstGeom@prst`: `wedgeRoundRectCallout`");
    expect(markdownFile.markdown).toContain("- `a:t#text`: `角四角`");
    expect(markdownFile.markdown).toContain("### Shape 002 (K3)");
    expect(markdownFile.markdown).toContain("- `a:prstGeom@prst`: `wedgeEllipseCallout`");
    expect(markdownFile.markdown).toContain("- `a:t#text`: `楕円`");
    expect(markdownFile.markdown).toContain("### Shape 003 (N3)");
    expect(markdownFile.markdown).toContain("- `a:prstGeom@prst`: `cloudCallout`");
    expect(markdownFile.markdown).toContain("- `a:t#text`: `雲`");
    expect(markdownFile.markdown).toContain("### Shape 004 (H8)");
    expect(markdownFile.markdown).toContain("- `a:prstGeom@prst`: `accentBorderCallout1`");
    expect(markdownFile.markdown).toContain("- `a:t#text`: `注釈`");
    expect(markdownFile.markdown).not.toContain("![shape_001.svg]");
    expect(markdownFile.markdown).not.toContain("## Images");
    expect(markdownFile.markdown).not.toContain("## Charts");
  });

  it("exports only markdown for the callout fixture archive when no SVG assets are generated", async () => {
    const api = bootCore();
    const fixtureName = "shape-callout-sample01.xlsx";
    const fixturePath = path.resolve(fixtureDir, "shape", fixtureName);
    const fixtureBytes = readFileSync(fixturePath);
    const arrayBuffer = fixtureBytes.buffer.slice(
      fixtureBytes.byteOffset,
      fixtureBytes.byteOffset + fixtureBytes.byteLength
    );

    const workbook = await api.parseWorkbook(arrayBuffer, fixtureName);
    const files = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true
    });
    const archive = api.createWorkbookExportArchive(workbook, files);
    const extracted = await api.unzipEntries(
      archive.buffer.slice(archive.byteOffset, archive.byteOffset + archive.byteLength)
    );

    expect(extracted.has("output/shape-callout-sample01.md")).toBe(true);
    expect(Array.from(extracted.keys())).toEqual(["output/shape-callout-sample01.md"]);

    const markdownText = new TextDecoder().decode(extracted.get("output/shape-callout-sample01.md"));
    expect(markdownText).toContain("## Shapes");
    expect(markdownText).toContain("- `a:prstGeom@prst`: `cloudCallout`");
    expect(markdownText).not.toContain("![shape_001.svg]");
  });

  it("parses the formula-shared fixture workbook with concrete shared-formula expectations", async () => {
    const api = bootCore();
    const fixtureName = "formula-shared-sample01.xlsx";
    const fixturePath = path.resolve(fixtureDir, "formula", fixtureName);
    const fixtureBytes = readFileSync(fixturePath);
    const arrayBuffer = fixtureBytes.buffer.slice(
      fixtureBytes.byteOffset,
      fixtureBytes.byteOffset + fixtureBytes.byteLength
    );

    const workbook = await api.parseWorkbook(arrayBuffer, fixtureName);
    const files = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true
    });
    const sheet = workbook.sheets[0];
    const markdownFile = files[0];

    expect(workbook.sheets).toHaveLength(1);
    expect(sheet.name).toBe("formula");
    expect(sheet.maxRow).toBe(13);
    expect(sheet.maxCol).toBe(4);
    expect(sheet.cells).toHaveLength(27);
    expect(sheet.cells.find((cell) => cell.address === "A1")?.outputValue).toBe("No");
    expect(sheet.cells.find((cell) => cell.address === "B1")?.outputValue).toBe("連番");
    expect(sheet.cells.find((cell) => cell.address === "D1")?.outputValue).toBe("shared formula サンプル");
    expect(sheet.cells.find((cell) => cell.address === "B3")?.formulaText).toBe("=B2+1");
    expect(sheet.cells.find((cell) => cell.address === "B3")?.outputValue).toBe("2");
    expect(sheet.cells.find((cell) => cell.address === "B3")?.resolutionStatus).toBe("resolved");
    expect(sheet.cells.find((cell) => cell.address === "B4")?.formulaText).toBe("=B3+1");
    expect(sheet.cells.find((cell) => cell.address === "B4")?.outputValue).toBe("3");
    expect(sheet.cells.find((cell) => cell.address === "B5")?.formulaText).toBe("=B4+1");
    expect(sheet.cells.find((cell) => cell.address === "B5")?.outputValue).toBe("4");
    expect(sheet.cells.find((cell) => cell.address === "B6")?.formulaText).toBe("=B5+1");
    expect(sheet.cells.find((cell) => cell.address === "B6")?.outputValue).toBe("5");
    expect(sheet.cells.find((cell) => cell.address === "B11")?.formulaText).toBe("=B10+1");
    expect(sheet.cells.find((cell) => cell.address === "B11")?.outputValue).toBe("10");
    expect(sheet.cells.find((cell) => cell.address === "B11")?.resolutionStatus).toBe("resolved");

    expect(markdownFile.fileName).toBe("formula-shared-sample01_001_formula.md");
    expect(markdownFile.summary.tables).toBe(1);
    expect(markdownFile.summary.merges).toBe(0);
    expect(markdownFile.summary.images).toBe(0);
    expect(markdownFile.summary.formulaDiagnostics).toHaveLength(9);
    expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual(["A1-B11"]);
    expect(markdownFile.markdown).toContain("# formula");
    expect(markdownFile.markdown).toContain("Workbook: formula-shared-sample01.xlsx");
    expect(markdownFile.markdown).toContain("| No | 連番 |");
    expect(markdownFile.markdown).toContain("| 1 | 1 |");
    expect(markdownFile.markdown).toContain("| 2 | 2 |");
    expect(markdownFile.markdown).toContain("| 5 | 5 |");
    expect(markdownFile.markdown).toContain("| 10 | 10 |");
  });

  it("parses the named-range fixture workbook with concrete defined-name expectations", async () => {
    const api = bootCore();
    const fixtureName = "named-range-sample01.xlsx";
    const fixturePath = path.resolve(fixtureDir, "named-range", fixtureName);
    const fixtureBytes = readFileSync(fixturePath);
    const arrayBuffer = fixtureBytes.buffer.slice(
      fixtureBytes.byteOffset,
      fixtureBytes.byteOffset + fixtureBytes.byteLength
    );

    const workbook = await api.parseWorkbook(arrayBuffer, fixtureName);
    const files = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true
    });
    const summarySheet = workbook.sheets[0];
    const otherSheet = workbook.sheets[1];
    const summaryFile = files[0];
    const otherFile = files[1];

    expect(workbook.definedNames).toEqual([
      {
        name: "BaseName",
        formulaText: "=Summary!$B$3",
        localSheetName: null
      },
      {
        name: "BaseRange",
        formulaText: "=Summary!$B$4:$B$5",
        localSheetName: null
      },
      {
        name: "LocalCross",
        formulaText: "=Other!$B$2",
        localSheetName: "Other"
      }
    ]);

    expect(workbook.sheets).toHaveLength(2);

    expect(summarySheet.name).toBe("Summary");
    expect(summarySheet.maxRow).toBe(13);
    expect(summarySheet.maxCol).toBe(4);
    expect(summarySheet.cells).toHaveLength(25);
    expect(summarySheet.cells.find((cell) => cell.address === "A1")?.outputValue).toBe("definedNames サンプル");
    expect(summarySheet.cells.find((cell) => cell.address === "B3")?.outputValue).toBe("Base");
    expect(summarySheet.cells.find((cell) => cell.address === "D3")?.formulaText).toBe("=BaseName");
    expect(summarySheet.cells.find((cell) => cell.address === "D3")?.outputValue).toBe("Base");
    expect(summarySheet.cells.find((cell) => cell.address === "D3")?.resolutionStatus).toBe("resolved");
    expect(summarySheet.cells.find((cell) => cell.address === "D4")?.formulaText).toBe("=SUM(BaseRange)");
    expect(summarySheet.cells.find((cell) => cell.address === "D4")?.outputValue).toBe("30");
    expect(summarySheet.cells.find((cell) => cell.address === "D4")?.resolutionStatus).toBe("resolved");

    expect(otherSheet.name).toBe("Other");
    expect(otherSheet.maxRow).toBe(2);
    expect(otherSheet.maxCol).toBe(4);
    expect(otherSheet.cells).toHaveLength(3);
    expect(otherSheet.cells.find((cell) => cell.address === "A1")?.outputValue).toBe("LocalCross元");
    expect(otherSheet.cells.find((cell) => cell.address === "B2")?.outputValue).toBe("CrossRef");
    expect(otherSheet.cells.find((cell) => cell.address === "D2")?.formulaText).toBe("=LocalCross");
    expect(otherSheet.cells.find((cell) => cell.address === "D2")?.outputValue).toBe("CrossRef");
    expect(otherSheet.cells.find((cell) => cell.address === "D2")?.resolutionStatus).toBe("resolved");

    expect(summaryFile.fileName).toBe("named-range-sample01_001_Summary.md");
    expect(summaryFile.summary.tables).toBe(1);
    expect(summaryFile.summary.tableScores.map((detail) => detail.range)).toEqual(["A3-B5"]);
    expect(summaryFile.summary.formulaDiagnostics).toHaveLength(2);
    expect(summaryFile.summary.formulaDiagnostics.every((diagnostic) => diagnostic.source === "cached_value")).toBe(true);
    expect(summaryFile.markdown).toContain("# Summary");
    expect(summaryFile.markdown).toContain("Workbook: named-range-sample01.xlsx");
    expect(summaryFile.markdown).toContain("| BaseName元 | Base |");
    expect(summaryFile.markdown).toContain("| BaseRange1 | 10 |");
    expect(summaryFile.markdown).toContain("| BaseRange2 | 20 |");

    expect(otherFile.fileName).toBe("named-range-sample01_002_Other.md");
    expect(otherFile.summary.tables).toBe(0);
    expect(otherFile.summary.tableScores).toHaveLength(0);
    expect(otherFile.summary.formulaDiagnostics).toHaveLength(1);
    expect(otherFile.markdown).toContain("# Other");
    expect(otherFile.markdown).toContain("LocalCross元");
    expect(otherFile.markdown).toContain("CrossRef");
  });

  it("parses the narrative fixture workbook with concrete narrative and table expectations", async () => {
    const api = bootCore();
    const fixtureName = "narrative-vs-table-sample01.xlsx";
    const fixturePath = path.resolve(fixtureDir, "narrative", fixtureName);
    const fixtureBytes = readFileSync(fixturePath);
    const arrayBuffer = fixtureBytes.buffer.slice(
      fixtureBytes.byteOffset,
      fixtureBytes.byteOffset + fixtureBytes.byteLength
    );

    const workbook = await api.parseWorkbook(arrayBuffer, fixtureName);
    const files = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true
    });
    const sheet = workbook.sheets[0];
    const markdownFile = files[0];

    expect(workbook.sheets).toHaveLength(1);
    expect(sheet.name).toBe("narrative-vs-table");
    expect(sheet.maxRow).toBe(13);
    expect(sheet.maxCol).toBe(6);
    expect(sheet.cells).toHaveLength(26);
    expect(sheet.cells.find((cell) => cell.address === "A1")?.outputValue).toBe("地の文と表の判定");
    expect(sheet.cells.find((cell) => cell.address === "A3")?.outputValue).toBe("この設計書は受注入力画面を説明する。");
    expect(sheet.cells.find((cell) => cell.address === "A4")?.outputValue).toBe("外部システムとの連携条件を以下に示す。");
    expect(sheet.cells.find((cell) => cell.address === "A5")?.outputValue).toBe("本文は罫線なしのままにする。");
    expect(sheet.cells.find((cell) => cell.address === "A7")?.outputValue).toBe("項目一覧");
    expect(sheet.cells.find((cell) => cell.address === "B8")?.outputValue).toBe("項番");
    expect(sheet.cells.find((cell) => cell.address === "F8")?.outputValue).toBe("備考");
    expect(sheet.cells.find((cell) => cell.address === "B10")?.formulaText).toBe("=B9+1");
    expect(sheet.cells.find((cell) => cell.address === "B10")?.outputValue).toBe("2");
    expect(sheet.cells.find((cell) => cell.address === "B11")?.formulaText).toBe("=B10+1");
    expect(sheet.cells.find((cell) => cell.address === "B11")?.outputValue).toBe("3");
    expect(sheet.cells.find((cell) => cell.address === "E11")?.outputValue).toBe("3月13日");
    expect(sheet.cells.find((cell) => cell.address === "B13")?.outputValue).toBe("※注記: この表はサンプルです。");

    expect(markdownFile.fileName).toBe("narrative-vs-table-sample01_001_narrative-vs-table.md");
    expect(markdownFile.summary.tables).toBe(1);
    expect(markdownFile.summary.merges).toBe(0);
    expect(markdownFile.summary.images).toBe(0);
    expect(markdownFile.summary.formulaDiagnostics).toHaveLength(2);
    expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual(["B8-F11"]);
    expect(markdownFile.markdown).toContain("# narrative-vs-table");
    expect(markdownFile.markdown).toContain("Workbook: narrative-vs-table-sample01.xlsx");
    expect(markdownFile.markdown).toContain("地の文と表の判定");
    expect(markdownFile.markdown).toContain("この設計書は受注入力画面を説明する。");
    expect(markdownFile.markdown).toContain("外部システムとの連携条件を以下に示す。");
    expect(markdownFile.markdown).toContain("本文は罫線なしのままにする。");
    expect(markdownFile.markdown).toContain("項目一覧");
    expect(markdownFile.markdown).toContain("### Table 001 (B8-F11)");
    expect(markdownFile.markdown).toContain("| 項番 | 項目名称 | 物理名 | 初期値 | 備考 |");
    expect(markdownFile.markdown).toContain("| 1 | コード | code | 101 | 何かのコード |");
    expect(markdownFile.markdown).toContain("| 3 | 登録日 | entrydate | 3月13日 | 何かの登録日 |");
    expect(markdownFile.markdown).toContain("※注記: この表はサンプルです。");
  });

  it("parses the edge-empty fixture workbook with concrete sparse-sheet expectations", async () => {
    const api = bootCore();
    const fixtureName = "edge-empty-sample01.xlsx";
    const fixturePath = path.resolve(fixtureDir, "edge", fixtureName);
    const fixtureBytes = readFileSync(fixturePath);
    const arrayBuffer = fixtureBytes.buffer.slice(
      fixtureBytes.byteOffset,
      fixtureBytes.byteOffset + fixtureBytes.byteLength
    );

    const workbook = await api.parseWorkbook(arrayBuffer, fixtureName);
    const files = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true
    });
    const sheet = workbook.sheets[0];
    const markdownFile = files[0];

    expect(workbook.sheets).toHaveLength(1);
    expect(sheet.name).toBe("edge-empty");
    expect(sheet.maxRow).toBe(7);
    expect(sheet.maxCol).toBe(3);
    expect(sheet.cells).toHaveLength(2);
    expect(sheet.cells.find((cell) => cell.address === "A1")?.outputValue).toBe("空系境界サンプル");
    expect(sheet.cells.find((cell) => cell.address === "C7")?.outputValue).toBe("only-value");

    expect(markdownFile.fileName).toBe("edge-empty-sample01_001_edge-empty.md");
    expect(markdownFile.summary.tables).toBe(0);
    expect(markdownFile.summary.merges).toBe(0);
    expect(markdownFile.summary.images).toBe(0);
    expect(markdownFile.summary.tableScores).toHaveLength(0);
    expect(markdownFile.markdown).toContain("# edge-empty");
    expect(markdownFile.markdown).toContain("Workbook: edge-empty-sample01.xlsx");
    expect(markdownFile.markdown).toContain("空系境界サンプル");
    expect(markdownFile.markdown).toContain("only-value");
    expect(markdownFile.markdown).not.toContain("### Table");
  });

  it("parses the edge-weird-sheetname fixture workbook with concrete sheet-name expectations", async () => {
    const api = bootCore();
    const fixtureName = "edge-weird-sheetname-sample01.xlsx";
    const fixturePath = path.resolve(fixtureDir, "edge", fixtureName);
    const fixtureBytes = readFileSync(fixturePath);
    const arrayBuffer = fixtureBytes.buffer.slice(
      fixtureBytes.byteOffset,
      fixtureBytes.byteOffset + fixtureBytes.byteLength
    );

    const workbook = await api.parseWorkbook(arrayBuffer, fixtureName);
    const files = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true
    });
    const sheet = workbook.sheets[0];
    const markdownFile = files[0];

    expect(workbook.sheets).toHaveLength(1);
    expect(sheet.name).toBe("A B-東京&大阪.01");
    expect(sheet.maxRow).toBe(4);
    expect(sheet.maxCol).toBe(4);
    expect(sheet.cells).toHaveLength(16);
    expect(sheet.cells.find((cell) => cell.address === "A1")?.outputValue).toBe("項番");
    expect(sheet.cells.find((cell) => cell.address === "B1")?.outputValue).toBe("名称");
    expect(sheet.cells.find((cell) => cell.address === "C4")?.outputValue).toBe("3月13日");
    expect(sheet.cells.find((cell) => cell.address === "D4")?.outputValue).toBe("何かの登録日");

    expect(markdownFile.fileName).toBe("edge-weird-sheetname-sample01_001_A_B-東京_大阪.01.md");
    expect(markdownFile.summary.tables).toBe(1);
    expect(markdownFile.summary.merges).toBe(0);
    expect(markdownFile.summary.images).toBe(0);
    expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual(["A1-D4"]);
    expect(markdownFile.markdown).toContain("# A B-東京&大阪.01");
    expect(markdownFile.markdown).toContain("Workbook: edge-weird-sheetname-sample01.xlsx");
    expect(markdownFile.markdown).toContain("Sheet: A B-東京&大阪.01");
    expect(markdownFile.markdown).toContain("### Table 001 (A1-D4)");
    expect(markdownFile.markdown).toContain("| 項番 | 名称 | 値 | 備考 |");
    expect(markdownFile.markdown).toContain("| 1 | コード | 101 | 何かのコード |");
    expect(markdownFile.markdown).toContain("| 3 | 登録日 | 3月13日 | 何かの登録日 |");
  });

  it("parses the table-basic-sample01 fixture workbook as two vertically adjacent independent tables", async () => {
    const api = bootCore();
    const fixtureName = "table-basic-sample01.xlsx";
    const fixturePath = path.resolve(fixtureDir, "table", fixtureName);
    const fixtureBytes = readFileSync(fixturePath);
    const arrayBuffer = fixtureBytes.buffer.slice(
      fixtureBytes.byteOffset,
      fixtureBytes.byteOffset + fixtureBytes.byteLength
    );

    const workbook = await api.parseWorkbook(arrayBuffer, fixtureName);
    const files = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true
    });
    const sheet = workbook.sheets[0];
    const markdownFile = files[0];

    expect(workbook.sheets).toHaveLength(1);
    expect(sheet.name).toBe("table-basic");
    expect(sheet.maxRow).toBe(13);
    expect(sheet.maxCol).toBe(6);
    expect(sheet.cells).toHaveLength(53);
    expect(sheet.cells.find((cell) => cell.address === "A1")?.outputValue).toBe("隣接するテーブルのテスト（縦に密接）");
    expect(sheet.cells.find((cell) => cell.address === "B2")?.outputValue).toBe("隣接するテーブルその1");
    expect(sheet.cells.find((cell) => cell.address === "B8")?.outputValue).toBe("隣接するテーブルその2");
    expect(sheet.cells.find((cell) => cell.address === "B5")?.formulaText).toBe("=B4+1");
    expect(sheet.cells.find((cell) => cell.address === "B13")?.formulaText).toBe("=B12+1");
    expect(sheet.cells.find((cell) => cell.address === "E12")?.outputValue).toBe("3月15日");
    expect(sheet.cells.find((cell) => cell.address === "F13")?.outputValue).toBe("更新した日");

    expect(markdownFile.fileName).toBe("table-basic-sample01_001_table-basic.md");
    expect(markdownFile.summary.tables).toBe(2);
    expect(markdownFile.summary.merges).toBe(0);
    expect(markdownFile.summary.images).toBe(0);
    expect(markdownFile.summary.formulaDiagnostics).toHaveLength(6);
    expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual([
      "B3-F7",
      "B9-F13"
    ]);
    expect(markdownFile.markdown).toContain("# table-basic");
    expect(markdownFile.markdown).toContain("Workbook: table-basic-sample01.xlsx");
    expect(markdownFile.markdown).toContain("隣接するテーブルのテスト（縦に密接）");
    expect(markdownFile.markdown).toContain("隣接するテーブルその1");
    expect(markdownFile.markdown).toContain("隣接するテーブルその2");
    expect(markdownFile.markdown).toContain("### Table 001 (B3-F7)");
    expect(markdownFile.markdown).toContain("### Table 002 (B9-F13)");
    expect(markdownFile.markdown).toContain("| 1 | コード | code | 101 | 何かのコード |");
    expect(markdownFile.markdown).toContain("| 3 | 登録日 | createdAt | 3月15日 | 登録した日 |");
  });

  it("parses the table-basic-sample02 fixture workbook as two horizontally adjacent independent tables", async () => {
    const api = bootCore();
    const fixtureName = "table-basic-sample02.xlsx";
    const fixturePath = path.resolve(fixtureDir, "table", fixtureName);
    const fixtureBytes = readFileSync(fixturePath);
    const arrayBuffer = fixtureBytes.buffer.slice(
      fixtureBytes.byteOffset,
      fixtureBytes.byteOffset + fixtureBytes.byteLength
    );

    const workbook = await api.parseWorkbook(arrayBuffer, fixtureName);
    const files = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true
    });
    const sheet = workbook.sheets[0];
    const markdownFile = files[0];

    expect(workbook.sheets).toHaveLength(1);
    expect(sheet.name).toBe("table-basic");
    expect(sheet.maxRow).toBe(7);
    expect(sheet.maxCol).toBe(12);
    expect(sheet.cells).toHaveLength(56);
    expect(sheet.cells.find((cell) => cell.address === "A1")?.outputValue).toBe("隣接するテーブルのテスト（横に密接）");
    expect(sheet.cells.find((cell) => cell.address === "B2")?.outputValue).toBe("隣接するテーブルその1");
    expect(sheet.cells.find((cell) => cell.address === "H2")?.outputValue).toBe("隣接するテーブルその2");
    expect(sheet.cells.find((cell) => cell.address === "G4")?.outputValue).toBe("確認");
    expect(sheet.cells.find((cell) => cell.address === "G6")?.outputValue).toBe("日付");
    expect(sheet.cells.find((cell) => cell.address === "H7")?.formulaText).toBe("=H6+1");
    expect(sheet.cells.find((cell) => cell.address === "K6")?.outputValue).toBe("3月15日");

    expect(markdownFile.fileName).toBe("table-basic-sample02_001_table-basic.md");
    expect(markdownFile.summary.tables).toBe(2);
    expect(markdownFile.summary.merges).toBe(0);
    expect(markdownFile.summary.images).toBe(0);
    expect(markdownFile.summary.formulaDiagnostics).toHaveLength(6);
    expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual([
      "B3-F7",
      "H3-L7"
    ]);
    expect(markdownFile.markdown).toContain("# table-basic");
    expect(markdownFile.markdown).toContain("Workbook: table-basic-sample02.xlsx");
    expect(markdownFile.markdown).toContain("隣接するテーブルのテスト（横に密接）");
    expect(markdownFile.markdown).toContain("### Table 001 (B3-F7)");
    expect(markdownFile.markdown).toContain("### Table 002 (H3-L7)");
    expect(markdownFile.markdown).toContain("| 1 | コード | code | 101 | 何かのコード |");
    expect(markdownFile.markdown).toContain("| 2 | 別名 | altname | Hanako | 何かの別名 |");
    expect(markdownFile.markdown).toContain("確認");
    expect(markdownFile.markdown).toContain("日付");
  });

  it("parses the table-basic-sample03 fixture workbook as four tightly adjacent independent tables", async () => {
    const api = bootCore();
    const fixtureName = "table-basic-sample03.xlsx";
    const fixturePath = path.resolve(fixtureDir, "table", fixtureName);
    const fixtureBytes = readFileSync(fixturePath);
    const arrayBuffer = fixtureBytes.buffer.slice(
      fixtureBytes.byteOffset,
      fixtureBytes.byteOffset + fixtureBytes.byteLength
    );

    const workbook = await api.parseWorkbook(arrayBuffer, fixtureName);
    const files = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true
    });
    const sheet = workbook.sheets[0];
    const markdownFile = files[0];

    expect(workbook.sheets).toHaveLength(1);
    expect(sheet.name).toBe("table-basic");
    expect(sheet.maxRow).toBe(13);
    expect(sheet.maxCol).toBe(12);
    expect(sheet.cells).toHaveLength(111);
    expect(sheet.cells.find((cell) => cell.address === "A1")?.outputValue).toBe("隣接するテーブルのテスト（縦横に密接）");
    expect(sheet.cells.find((cell) => cell.address === "B8")?.outputValue).toBe("隣接するテーブルその3");
    expect(sheet.cells.find((cell) => cell.address === "H8")?.outputValue).toBe("隣接するテーブルその4");
    expect(sheet.cells.find((cell) => cell.address === "G10")?.outputValue).toBe("確認");
    expect(sheet.cells.find((cell) => cell.address === "H13")?.formulaText).toBe("=H12+1");
    expect(sheet.cells.find((cell) => cell.address === "K13")?.outputValue).toBe("3月19日");
    expect(sheet.cells.find((cell) => cell.address === "L11")?.outputValue).toBe("何かの別名");

    expect(markdownFile.fileName).toBe("table-basic-sample03_001_table-basic.md");
    expect(markdownFile.summary.tables).toBe(4);
    expect(markdownFile.summary.merges).toBe(0);
    expect(markdownFile.summary.images).toBe(0);
    expect(markdownFile.summary.formulaDiagnostics).toHaveLength(12);
    expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual([
      "B3-F7",
      "H3-L7",
      "B9-F13",
      "H9-L13"
    ]);
    expect(markdownFile.markdown).toContain("# table-basic");
    expect(markdownFile.markdown).toContain("Workbook: table-basic-sample03.xlsx");
    expect(markdownFile.markdown).toContain("### Table 001 (B3-F7)");
    expect(markdownFile.markdown).toContain("### Table 002 (H3-L7)");
    expect(markdownFile.markdown).toContain("### Table 003 (B9-F13)");
    expect(markdownFile.markdown).toContain("### Table 004 (H9-L13)");
    expect(markdownFile.markdown).not.toContain("### Table 005");
    expect(markdownFile.markdown).not.toContain("### Table 002 (B3-L13)");
    expect(markdownFile.markdown).toContain("| 1 | コード | code | 301 | 何かのコード |");
    expect(markdownFile.markdown).toContain("| 2 | 別名 | altname | Sawada | 何かの別名 |");
  });

  it("parses the table-basic-sample11 fixture workbook as one merge-heavy graph-paper table", async () => {
    const api = bootCore();
    const fixtureName = "table-basic-sample11.xlsx";
    const fixturePath = path.resolve(fixtureDir, "table", fixtureName);
    const fixtureBytes = readFileSync(fixturePath);
    const arrayBuffer = fixtureBytes.buffer.slice(
      fixtureBytes.byteOffset,
      fixtureBytes.byteOffset + fixtureBytes.byteLength
    );

    const workbook = await api.parseWorkbook(arrayBuffer, fixtureName);
    const files = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true
    });
    const sheet = workbook.sheets[0];
    const markdownFile = files[0];

    expect(workbook.sheets).toHaveLength(1);
    expect(sheet.name).toBe("table-basic");
    expect(sheet.maxRow).toBe(7);
    expect(sheet.maxCol).toBe(20);
    expect(sheet.cells).toHaveLength(97);
    expect(sheet.merges).toHaveLength(20);
    expect(sheet.cells.find((cell) => cell.address === "A1")?.outputValue).toBe("方眼紙的様式のテスト");
    expect(sheet.cells.find((cell) => cell.address === "B2")?.outputValue).toBe("テーブルその1");
    expect(sheet.cells.find((cell) => cell.address === "B5")?.formulaText).toBe("=B4+1");
    expect(sheet.cells.find((cell) => cell.address === "P4")?.outputValue).toBe("何かのコード");
    expect(sheet.cells.find((cell) => cell.address === "L6")?.outputValue).toBe("3月13日");

    expect(markdownFile.fileName).toBe("table-basic-sample11_001_table-basic.md");
    expect(markdownFile.summary.tables).toBe(1);
    expect(markdownFile.summary.merges).toBe(20);
    expect(markdownFile.summary.images).toBe(0);
    expect(markdownFile.summary.formulaDiagnostics).toHaveLength(3);
    expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual(["B3-T7"]);
    expect(markdownFile.markdown).toContain("# table-basic");
    expect(markdownFile.markdown).toContain("Workbook: table-basic-sample11.xlsx");
    expect(markdownFile.markdown).toContain("方眼紙的様式のテスト");
    expect(markdownFile.markdown).toContain("### Table 001 (B3-T7)");
    expect(markdownFile.markdown).toContain("| 1 | コード | code | 101 | 何かのコード |");
    expect(markdownFile.markdown).toContain("| 4 | 更新日 | updatedate | 3月14日 | 何かの更新日 |");
  });

  it("parses the table-basic-sample12 fixture workbook as two vertically separated merge-heavy graph-paper tables", async () => {
    const api = bootCore();
    const fixtureName = "table-basic-sample12.xlsx";
    const fixturePath = path.resolve(fixtureDir, "table", fixtureName);
    const fixtureBytes = readFileSync(fixturePath);
    const arrayBuffer = fixtureBytes.buffer.slice(
      fixtureBytes.byteOffset,
      fixtureBytes.byteOffset + fixtureBytes.byteLength
    );

    const workbook = await api.parseWorkbook(arrayBuffer, fixtureName);
    const files = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true
    });
    const sheet = workbook.sheets[0];
    const markdownFile = files[0];

    expect(workbook.sheets).toHaveLength(1);
    expect(sheet.name).toBe("table-basic");
    expect(sheet.maxRow).toBe(14);
    expect(sheet.maxCol).toBe(20);
    expect(sheet.cells).toHaveLength(194);
    expect(sheet.merges).toHaveLength(40);
    expect(sheet.cells.find((cell) => cell.address === "A1")?.outputValue).toBe("方眼紙的様式のテスト");
    expect(sheet.cells.find((cell) => cell.address === "B8")?.outputValue).toBe("方眼紙風のためにセル結合が多用されます");
    expect(sheet.cells.find((cell) => cell.address === "B9")?.outputValue).toBe("テーブルその2");
    expect(sheet.cells.find((cell) => cell.address === "B14")?.formulaText).toBe("=B13+1");
    expect(sheet.cells.find((cell) => cell.address === "L13")?.outputValue).toBe("3月13日");

    expect(markdownFile.fileName).toBe("table-basic-sample12_001_table-basic.md");
    expect(markdownFile.summary.tables).toBe(2);
    expect(markdownFile.summary.merges).toBe(40);
    expect(markdownFile.summary.images).toBe(0);
    expect(markdownFile.summary.formulaDiagnostics).toHaveLength(6);
    expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual([
      "B3-T7",
      "B10-T14"
    ]);
    expect(markdownFile.markdown).toContain("Workbook: table-basic-sample12.xlsx");
    expect(markdownFile.markdown).toContain("### Table 001 (B3-T7)");
    expect(markdownFile.markdown).toContain("### Table 002 (B10-T14)");
    expect(markdownFile.markdown).toContain("方眼紙風のためにセル結合が多用されます");
    expect(markdownFile.markdown).toContain("テーブルその2");
  });

  it("parses the table-basic-sample13 fixture workbook as four merge-heavy graph-paper tables", async () => {
    const api = bootCore();
    const fixtureName = "table-basic-sample13.xlsx";
    const fixturePath = path.resolve(fixtureDir, "table", fixtureName);
    const fixtureBytes = readFileSync(fixturePath);
    const arrayBuffer = fixtureBytes.buffer.slice(
      fixtureBytes.byteOffset,
      fixtureBytes.byteOffset + fixtureBytes.byteLength
    );

    const workbook = await api.parseWorkbook(arrayBuffer, fixtureName);
    const files = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true
    });
    const sheet = workbook.sheets[0];
    const markdownFile = files[0];

    expect(workbook.sheets).toHaveLength(1);
    expect(sheet.name).toBe("table-basic");
    expect(sheet.maxRow).toBe(14);
    expect(sheet.maxCol).toBe(40);
    expect(sheet.cells).toHaveLength(389);
    expect(sheet.merges).toHaveLength(80);
    expect(sheet.cells.find((cell) => cell.address === "B9")?.outputValue).toBe("テーブルその3");
    expect(sheet.cells.find((cell) => cell.address === "V9")?.outputValue).toBe("テーブルその4");
    expect(sheet.cells.find((cell) => cell.address === "V14")?.formulaText).toBe("=V13+1");
    expect(sheet.cells.find((cell) => cell.address === "AJ11")?.outputValue).toBe("何かのコード");
    expect(sheet.cells.find((cell) => cell.address === "AF12")?.outputValue).toBe("Sabro");

    expect(markdownFile.fileName).toBe("table-basic-sample13_001_table-basic.md");
    expect(markdownFile.summary.tables).toBe(4);
    expect(markdownFile.summary.merges).toBe(80);
    expect(markdownFile.summary.images).toBe(0);
    expect(markdownFile.summary.formulaDiagnostics).toHaveLength(12);
    expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual([
      "B3-T7",
      "V3-AN7",
      "B10-T14",
      "V10-AN14"
    ]);
    expect(markdownFile.markdown).toContain("Workbook: table-basic-sample13.xlsx");
    expect(markdownFile.markdown).toContain("### Table 001 (B3-T7)");
    expect(markdownFile.markdown).toContain("### Table 002 (V3-AN7)");
    expect(markdownFile.markdown).toContain("### Table 003 (B10-T14)");
    expect(markdownFile.markdown).toContain("### Table 004 (V10-AN14)");
    expect(markdownFile.markdown).toContain("| 1 | コード | code | 401 | 何かのコード |");
    expect(markdownFile.markdown).toContain("| 2 | 名前 | name | Sabro | 何かの名前 |");
    expect(markdownFile.markdown).toContain("| 2 | 名前 | name | Jiro | 何かの名前 |");
  });

  it("parses the table-basic-sample14 fixture workbook as one graph-paper table even with a few unmerged cells", async () => {
    const api = bootCore();
    const fixtureName = "table-basic-sample14.xlsx";
    const fixturePath = path.resolve(fixtureDir, "table", fixtureName);
    const fixtureBytes = readFileSync(fixturePath);
    const arrayBuffer = fixtureBytes.buffer.slice(
      fixtureBytes.byteOffset,
      fixtureBytes.byteOffset + fixtureBytes.byteLength
    );

    const workbook = await api.parseWorkbook(arrayBuffer, fixtureName);
    const files = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true
    });
    const sheet = workbook.sheets[0];
    const markdownFile = files[0];

    expect(workbook.sheets).toHaveLength(1);
    expect(sheet.name).toBe("table-basic");
    expect(sheet.maxRow).toBe(8);
    expect(sheet.maxCol).toBe(20);
    expect(sheet.cells).toHaveLength(98);
    expect(sheet.merges).toHaveLength(18);
    expect(sheet.cells.find((cell) => cell.address === "A1")?.outputValue).toBe("方眼紙的様式のテスト");
    expect(sheet.cells.find((cell) => cell.address === "B2")?.outputValue).toBe("テーブルその1");
    expect(sheet.cells.find((cell) => cell.address === "B8")?.outputValue).toContain("たまに結合漏れのセルがある場合");
    expect(sheet.cells.find((cell) => cell.address === "L5")?.outputValue).toBe("Taro");
    expect(sheet.cells.find((cell) => cell.address === "L6")?.outputValue).toBe("3月13日");

    expect(markdownFile.fileName).toBe("table-basic-sample14_001_table-basic.md");
    expect(markdownFile.summary.tables).toBe(1);
    expect(markdownFile.summary.merges).toBe(18);
    expect(markdownFile.summary.images).toBe(0);
    expect(markdownFile.summary.formulaDiagnostics).toHaveLength(3);
    expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual(["B3-T7"]);
    expect(markdownFile.markdown).toContain("Workbook: table-basic-sample14.xlsx");
    expect(markdownFile.markdown).toContain("### Table 001 (B3-T7)");
    expect(markdownFile.markdown).toContain("| 2 | 名前 | name | Taro | 何かの名前 |");
    expect(markdownFile.markdown).toContain("たまに結合漏れのセルがある場合");
  });

  it("parses the table-basic-sample15 fixture workbook as one graph-paper table with a vertical merged note", async () => {
    const api = bootCore();
    const fixtureName = "table-basic-sample15.xlsx";
    const fixturePath = path.resolve(fixtureDir, "table", fixtureName);
    const fixtureBytes = readFileSync(fixturePath);
    const arrayBuffer = fixtureBytes.buffer.slice(
      fixtureBytes.byteOffset,
      fixtureBytes.byteOffset + fixtureBytes.byteLength
    );

    const workbook = await api.parseWorkbook(arrayBuffer, fixtureName);
    const files = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true
    });
    const sheet = workbook.sheets[0];
    const markdownFile = files[0];

    expect(workbook.sheets).toHaveLength(1);
    expect(sheet.name).toBe("table-basic");
    expect(sheet.maxRow).toBe(8);
    expect(sheet.maxCol).toBe(20);
    expect(sheet.cells).toHaveLength(98);
    expect(sheet.merges).toHaveLength(19);
    expect(sheet.cells.find((cell) => cell.address === "B8")?.outputValue).toBe("※方眼紙＋結合＋さらに縦結合");
    expect(sheet.cells.find((cell) => cell.address === "P5")?.outputValue).toBe("何かの名前");
    expect(sheet.cells.find((cell) => cell.address === "P6")?.outputValue).toBe("登録および更新日");
    expect(sheet.cells.find((cell) => cell.address === "P7")?.outputValue).toBe("");

    expect(markdownFile.fileName).toBe("table-basic-sample15_001_table-basic.md");
    expect(markdownFile.summary.tables).toBe(1);
    expect(markdownFile.summary.merges).toBe(19);
    expect(markdownFile.summary.images).toBe(0);
    expect(markdownFile.summary.formulaDiagnostics).toHaveLength(3);
    expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual(["B3-T7"]);
    expect(markdownFile.markdown).toContain("Workbook: table-basic-sample15.xlsx");
    expect(markdownFile.markdown).toContain("### Table 001 (B3-T7)");
    expect(markdownFile.markdown).toContain("| 3 | 登録日 | entrydate | 3月13日 | 登録および更新日 |");
    expect(markdownFile.markdown).toContain("| 4 | 更新日 | updatedate | 3月14日 | [MERGED↑] |");
    expect(markdownFile.markdown).toContain("※方眼紙＋結合＋さらに縦結合");
  });

  it("parses the table-basic-sample16 fixture workbook and keeps the extra value column caused by a merge gap", async () => {
    const api = bootCore();
    const fixtureName = "table-basic-sample16.xlsx";
    const fixturePath = path.resolve(fixtureDir, "table", fixtureName);
    const fixtureBytes = readFileSync(fixturePath);
    const arrayBuffer = fixtureBytes.buffer.slice(
      fixtureBytes.byteOffset,
      fixtureBytes.byteOffset + fixtureBytes.byteLength
    );

    const workbook = await api.parseWorkbook(arrayBuffer, fixtureName);
    const files = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true
    });
    const sheet = workbook.sheets[0];
    const markdownFile = files[0];

    expect(workbook.sheets).toHaveLength(1);
    expect(sheet.name).toBe("table-basic");
    expect(sheet.maxRow).toBe(8);
    expect(sheet.maxCol).toBe(20);
    expect(sheet.cells).toHaveLength(98);
    expect(sheet.merges).toHaveLength(18);
    expect(sheet.cells.find((cell) => cell.address === "B8")?.outputValue).toContain("たまに結合漏れのセルがあって");
    expect(sheet.cells.find((cell) => cell.address === "L5")?.outputValue).toBe("Taro");
    expect(sheet.cells.find((cell) => cell.address === "N5")?.outputValue).toBe("Ito");

    expect(markdownFile.fileName).toBe("table-basic-sample16_001_table-basic.md");
    expect(markdownFile.summary.tables).toBe(1);
    expect(markdownFile.summary.merges).toBe(18);
    expect(markdownFile.summary.images).toBe(0);
    expect(markdownFile.summary.formulaDiagnostics).toHaveLength(3);
    expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual(["B3-T7"]);
    expect(markdownFile.markdown).toContain("Workbook: table-basic-sample16.xlsx");
    expect(markdownFile.markdown).toContain("### Table 001 (B3-T7)");
    expect(markdownFile.markdown).toContain("| 項番 | 項目名称 | 物理名 | デフォルト値 | [MERGED←] | 備考 |");
    expect(markdownFile.markdown).toContain("| 2 | 名前 | name | Taro | Ito | 何かの名前 |");
    expect(markdownFile.markdown).toContain("たまに結合漏れのセルがあって、さらに複数文字が登場");
  });

  it("parses the table-border-priority-sample01 fixture workbook differently between balanced and border modes", async () => {
    const api = bootCore();
    const fixtureName = "table-border-priority-sample01.xlsx";
    const fixturePath = path.resolve(fixtureDir, "table", fixtureName);
    const fixtureBytes = readFileSync(fixturePath);
    const arrayBuffer = fixtureBytes.buffer.slice(
      fixtureBytes.byteOffset,
      fixtureBytes.byteOffset + fixtureBytes.byteLength
    );

    const workbook = await api.parseWorkbook(arrayBuffer, fixtureName);
    const balancedFiles = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true,
      tableDetectionMode: "balanced"
    });
    const borderFiles = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true,
      tableDetectionMode: "border"
    });
    const sheet = workbook.sheets[0];
    const balancedFile = balancedFiles[0];
    const borderFile = borderFiles[0];

    expect(workbook.sheets).toHaveLength(1);
    expect(sheet.name).toBe("border-priority");
    expect(sheet.maxRow).toBe(6);
    expect(sheet.maxCol).toBe(2);
    expect(sheet.cells).toHaveLength(6);
    expect(sheet.cells.find((cell) => cell.address === "A1")?.outputValue).toBe("border-priority fixture");
    expect(sheet.cells.find((cell) => cell.address === "A3")?.outputValue).toBe("項目");
    expect(sheet.cells.find((cell) => cell.address === "B4")?.outputValue).toBe("100");

    expect(balancedFile.fileName).toBe("table-border-priority-sample01_001_border-priority.md");
    expect(balancedFile.summary.tables).toBe(1);
    expect(balancedFile.summary.tableDetectionMode).toBe("balanced");
    expect(balancedFile.summary.tableScores.map((detail) => detail.range)).toEqual(["A3-B4"]);
    expect(balancedFile.markdown).toContain("Workbook: table-border-priority-sample01.xlsx");
    expect(balancedFile.markdown).toContain("### Table 001 (A3-B4)");
    expect(balancedFile.markdown).toContain("| 項目 | 値 |");
    expect(balancedFile.markdown).toContain("| A | 100 |");

    expect(borderFile.fileName).toBe("table-border-priority-sample01_001_border-priority.md");
    expect(borderFile.summary.tables).toBe(0);
    expect(borderFile.summary.tableDetectionMode).toBe("border");
    expect(borderFile.summary.tableScores).toHaveLength(0);
    expect(borderFile.markdown).toContain("Workbook: table-border-priority-sample01.xlsx");
    expect(borderFile.markdown).not.toContain("### Table 001");
    expect(borderFile.markdown).toContain("項目");
    expect(borderFile.markdown).toContain("100");
    expect(borderFile.markdown).toContain("※罫線優先モード確認用");
  });

  it("expands merged cells with structural tokens", () => {
    const api = bootCore();
    const matrix = [
      ["見出し", "", ""],
      ["", "", ""],
      ["値", "2", "3"]
    ];
    api.applyMergeTokens(matrix, [
      { startRow: 1, startCol: 1, endRow: 2, endCol: 3, ref: "A1:C2" }
    ], 1, 1, 3, 3);

    expect(matrix[0][1]).toBe("[MERGED←]");
    expect(matrix[0][2]).toBe("[MERGED←]");
    expect(matrix[1][0]).toBe("[MERGED↑]");
    expect(matrix[1][2]).toBe("[MERGED↑]");
  });

  it("extracts narrative blocks outside tables", () => {
    const api = bootCore();
    const sheet = {
      name: "Summary",
      index: 1,
      path: "xl/worksheets/sheet1.xml",
      merges: [],
      maxRow: 8,
      maxCol: 4,
      cells: [
        { row: 1, col: 1, outputValue: "このシステムは" },
        { row: 1, col: 2, outputValue: "受注を管理します。" },
        { row: 2, col: 1, outputValue: "外部IFと連携します。" },
        { row: 5, col: 1, outputValue: "ID" },
        { row: 5, col: 2, outputValue: "名称" },
        { row: 6, col: 1, outputValue: "1" },
        { row: 6, col: 2, outputValue: "販売" }
      ]
    };
    const tables = [{ startRow: 5, startCol: 1, endRow: 6, endCol: 2 }];

    const blocks = api.extractNarrativeBlocks({ name: "narrative.xlsx", sheets: [] }, sheet, tables);

    expect(blocks).toHaveLength(1);
    expect(blocks[0].lines.join("\n")).toContain("このシステムは 受注を管理します。");
    expect(blocks[0].lines.join("\n")).toContain("外部IFと連携します。");
  });

  it("renders indented narrative rows as heading plus markdown bullets", () => {
    const api = bootCore();
    const workbook = { name: "todo.xlsx" };
    const sheet = {
      name: "To Do リスト",
      index: 1,
      path: "xl/worksheets/sheet1.xml",
      merges: [],
      tables: [],
      images: [],
      maxRow: 8,
      maxCol: 2,
      cells: [
        { row: 1, col: 1, address: "A1", valueType: "str", rawValue: "To Do リスト", outputValue: "To Do リスト", formulaText: "", resolutionStatus: null, styleIndex: 0, borders: { top: false, bottom: false, left: false, right: false }, numFmtId: 0, formatCode: "General" },
        { row: 2, col: 1, address: "A2", valueType: "str", rawValue: "学校が始まる前に確認します", outputValue: "学校が始まる前に確認します", formulaText: "", resolutionStatus: null, styleIndex: 0, borders: { top: false, bottom: false, left: false, right: false }, numFmtId: 0, formatCode: "General" },
        { row: 3, col: 1, address: "A3", valueType: "str", rawValue: "完了 タスク", outputValue: "完了 タスク", formulaText: "", resolutionStatus: null, styleIndex: 0, borders: { top: false, bottom: false, left: false, right: false }, numFmtId: 0, formatCode: "General" },
        { row: 4, col: 2, address: "B4", valueType: "str", rawValue: "登録フォームに記入する", outputValue: "登録フォームに記入する", formulaText: "", resolutionStatus: null, styleIndex: 0, borders: { top: false, bottom: false, left: false, right: false }, numFmtId: 0, formatCode: "General" },
        { row: 5, col: 2, address: "B5", valueType: "str", rawValue: "健康診断をスケジュールする", outputValue: "健康診断をスケジュールする", formulaText: "", resolutionStatus: null, styleIndex: 0, borders: { top: false, bottom: false, left: false, right: false }, numFmtId: 0, formatCode: "General" },
        { row: 6, col: 2, address: "B6", valueType: "str", rawValue: "予防接種を確認する", outputValue: "予防接種を確認する", formulaText: "", resolutionStatus: null, styleIndex: 0, borders: { top: false, bottom: false, left: false, right: false }, numFmtId: 0, formatCode: "General" },
        { row: 7, col: 2, address: "B7", valueType: "str", rawValue: "教師に会う", outputValue: "教師に会う", formulaText: "", resolutionStatus: null, styleIndex: 0, borders: { top: false, bottom: false, left: false, right: false }, numFmtId: 0, formatCode: "General" }
      ]
    };

    const markdownFile = api.convertSheetToMarkdown(workbook, sheet, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true
    });

    expect(markdownFile.markdown).toContain("To Do リスト");
    expect(markdownFile.markdown).toContain("学校が始まる前に確認します");
    expect(markdownFile.markdown).toContain("### 完了 タスク");
    expect(markdownFile.markdown).toContain("- 登録フォームに記入する");
    expect(markdownFile.markdown).toContain("- 健康診断をスケジュールする");
    expect(markdownFile.markdown).toContain("- 予防接種を確認する");
    expect(markdownFile.markdown).toContain("- 教師に会う");
  });

  it("parses a minimal workbook and converts a sheet to markdown", async () => {
    const api = bootCore();
    const workbookXml = `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Summary" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`;
    const workbookRelsXml = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`;
    const sharedStringsXml = `<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="5" uniqueCount="5">
  <si><t>Region</t></si>
  <si><t>Sales</t></si>
  <si><t>East</t></si>
  <si><t>West</t></si>
  <si><t>Overview text</t></si>
</sst>`;
    const stylesXml = `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <borders count="3">
    <border><left/><right/><top/><bottom/></border>
    <border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/></border>
    <border><left/><right/><top/><bottom/></border>
  </borders>
  <cellXfs count="3">
    <xf borderId="0"/>
    <xf borderId="1"/>
    <xf borderId="2" numFmtId="14" applyNumberFormat="1"/>
  </cellXfs>
</styleSheet>`;
    const sheetXml = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>4</v></c>
    </row>
    <row r="3">
      <c r="A3" t="s" s="1"><v>0</v></c>
      <c r="B3" t="s" s="1"><v>1</v></c>
      <c r="C3" s="1"><v>100</v></c>
    </row>
    <row r="4">
      <c r="A4" t="s" s="1"><v>2</v></c>
      <c r="B4" t="s" s="1"><v>3</v></c>
      <c r="C4" s="1"><v>90</v></c>
    </row>
    <row r="6">
      <c r="A6" t="inlineStr"><is><t>Date</t></is></c>
      <c r="B6" s="2"><v>45292</v></c>
    </row>
  </sheetData>
</worksheet>`;
    const zip = createStoredZip([
      { name: "[Content_Types].xml", data: `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>` },
      { name: "xl/workbook.xml", data: workbookXml },
      { name: "xl/_rels/workbook.xml.rels", data: workbookRelsXml },
      { name: "xl/sharedStrings.xml", data: sharedStringsXml },
      { name: "xl/styles.xml", data: stylesXml },
      { name: "xl/worksheets/sheet1.xml", data: sheetXml }
    ]);

    const workbook = await api.parseWorkbook(zip, "sales.xlsx");
    const files = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true
    });
    const rawFiles = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true,
      outputMode: "raw"
    });
    const bothFiles = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true,
      outputMode: "both"
    });

    expect(workbook.sheets).toHaveLength(1);
    expect(workbook.sheets[0].name).toBe("Summary");
    expect(files).toHaveLength(1);
    expect(files[0].fileName).toBe("sales_001_Summary.md");
    expect(rawFiles[0].fileName).toBe("sales_001_Summary_raw.md");
    expect(bothFiles[0].fileName).toBe("sales_001_Summary_both.md");
    expect(files[0].markdown).toContain("# Summary");
    expect(files[0].markdown).toContain("Workbook: sales.xlsx");
    expect(files[0].markdown).toContain("Overview text");
    expect(files[0].markdown).toContain("| Region | Sales | 100 |");
    expect(files[0].markdown).toContain("| East | West | 90 |");
    expect(files[0].markdown).toContain("2024/1/1");
    expect(rawFiles[0].markdown).toContain("Date 45292");
    expect(bothFiles[0].markdown).toContain("Date 2024/1/1 [raw=45292]");
    expect(api.createSummaryText(bothFiles[0])).toContain("Output mode: both");
    expect(files[0].summary.images).toBe(0);
    expect(files[0].summary.tableScores).toHaveLength(1);
    expect(files[0].summary.tableScores[0].range).toBe("A3-C4");
    expect(files[0].summary.tableScores[0].score).toBeGreaterThanOrEqual(4);
    expect(files[0].summary.tableScores[0].reasons.join(" ")).toContain("Has borders");
  });

  it("extracts drawing image metadata and renders markdown image section", async () => {
    const api = bootCore();
    const workbookXml = `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Summary" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`;
    const workbookRelsXml = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`;
    const sheetXml = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Sheet with image</t></is></c>
    </row>
  </sheetData>
</worksheet>`;
    const sheetRelsXml = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`;
    const drawingXml = `<?xml version="1.0" encoding="UTF-8"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:oneCellAnchor>
    <xdr:from>
      <xdr:col>2</xdr:col>
      <xdr:row>1</xdr:row>
    </xdr:from>
    <xdr:pic>
      <xdr:blipFill>
        <a:blip r:embed="rIdImage1"/>
      </xdr:blipFill>
    </xdr:pic>
  </xdr:oneCellAnchor>
</xdr:wsDr>`;
    const drawingRelsXml = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImage1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/sample.png"/>
</Relationships>`;
    const pngBytes = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);
    const zip = createStoredZip([
      { name: "[Content_Types].xml", data: `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>` },
      { name: "xl/workbook.xml", data: workbookXml },
      { name: "xl/_rels/workbook.xml.rels", data: workbookRelsXml },
      { name: "xl/worksheets/sheet1.xml", data: sheetXml },
      { name: "xl/worksheets/_rels/sheet1.xml.rels", data: sheetRelsXml },
      { name: "xl/drawings/drawing1.xml", data: drawingXml },
      { name: "xl/drawings/_rels/drawing1.xml.rels", data: drawingRelsXml },
      { name: "xl/media/sample.png", data: pngBytes }
    ]);

    const workbook = await api.parseWorkbook(zip, "diagram.xlsx");
    const files = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true
    });

    expect(workbook.sheets[0].images).toHaveLength(1);
    expect(workbook.sheets[0].images[0].anchor).toBe("C2");
    expect(workbook.sheets[0].images[0].path).toBe("assets/Summary/image_001.png");
    expect(files[0].summary.images).toBe(1);
    expect(files[0].markdown).toContain("## Images");
    expect(files[0].markdown).toContain("### Image 001 (C2)");
    expect(files[0].markdown).toContain("![image_001.png](assets/Summary/image_001.png)");

    const archive = api.createWorkbookExportArchive(workbook, files);
    const extracted = await api.unzipEntries(archive.buffer.slice(archive.byteOffset, archive.byteOffset + archive.byteLength));
    expect(extracted.has("output/diagram.md")).toBe(true);
    expect(extracted.has("output/assets/Summary/image_001.png")).toBe(true);
    const markdownText = new TextDecoder().decode(extracted.get("output/diagram.md"));
    expect(markdownText).toContain("<!-- diagram_001_Summary -->");
    expect(markdownText).toContain("## Images");
    expect(extracted.get("output/assets/Summary/image_001.png")).toBeDefined();
  });

  it("resolves simple same-workbook formulas when cached value is missing", async () => {
    const api = bootCore();
    const workbookXml = `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <definedNames>
    <definedName name="BaseName">Summary!$A$1</definedName>
    <definedName name="BaseRange">Summary!$G$1:$H$1</definedName>
    <definedName name="LocalCross" localSheetId="0">'Other Sheet'!$B$2</definedName>
  </definedNames>
  <sheets>
    <sheet name="Summary" sheetId="1" r:id="rId1"/>
    <sheet name="Other Sheet" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>`;
    const workbookRelsXml = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>`;
    const sheet1Xml = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Base</t></is></c>
      <c r="B1"><f>A1</f></c>
      <c r="C1"><f>'Other Sheet'!B2</f></c>
      <c r="D1"><f>BaseName</f></c>
      <c r="E1"><f>SUM(BaseRange)</f></c>
      <c r="F1"><f>LocalCross</f></c>
      <c r="G1"><v>10</v></c>
      <c r="H1"><v>20</v></c>
    </row>
  </sheetData>
</worksheet>`;
    const sheet2Xml = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="2">
      <c r="B2" t="inlineStr"><is><t>CrossRef</t></is></c>
    </row>
  </sheetData>
</worksheet>`;
    const zip = createStoredZip([
      { name: "[Content_Types].xml", data: `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>` },
      { name: "xl/workbook.xml", data: workbookXml },
      { name: "xl/_rels/workbook.xml.rels", data: workbookRelsXml },
      { name: "xl/worksheets/sheet1.xml", data: sheet1Xml },
      { name: "xl/worksheets/sheet2.xml", data: sheet2Xml }
    ]);

    const workbook = await api.parseWorkbook(zip, "formula.xlsx");
    const summarySheet = workbook.sheets[0];
    const b1 = summarySheet.cells.find((cell) => cell.address === "B1");
    const c1 = summarySheet.cells.find((cell) => cell.address === "C1");
    const d1 = summarySheet.cells.find((cell) => cell.address === "D1");
    const e1 = summarySheet.cells.find((cell) => cell.address === "E1");
    const f1 = summarySheet.cells.find((cell) => cell.address === "F1");
    const files = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true
    });

    expect(b1?.outputValue).toBe("Base");
    expect(b1?.resolutionStatus).toBe("resolved");
    expect(c1?.outputValue).toBe("CrossRef");
    expect(c1?.resolutionStatus).toBe("resolved");
    expect(d1?.outputValue).toBe("Base");
    expect(e1?.outputValue).toBe("30");
    expect(f1?.outputValue).toBe("CrossRef");
    expect(files[0].summary.formulaDiagnostics).toHaveLength(5);
    expect(files[0].summary.formulaDiagnostics[0].status).toBe("resolved");
    expect(files[0].summary.formulaDiagnostics[4].status).toBe("resolved");
    expect(api.createSummaryText(files[0])).toContain("Formula resolved: 5");
  });

  it("evaluates simple arithmetic formulas over same-workbook references", async () => {
    const api = bootCore();
    const workbookXml = `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Summary" sheetId="1" r:id="rId1"/>
    <sheet name="Other" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>`;
    const workbookRelsXml = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>`;
    const sheet1Xml = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><v>10</v></c>
      <c r="B1"><v>5</v></c>
      <c r="C1"><f>A1+B1</f></c>
      <c r="D1"><f>(A1-B1)*2</f></c>
      <c r="E1"><f>Other!B2/2</f></c>
      <c r="F1"><f>$A$1+$B1</f></c>
      <c r="G1"><f>Other!$B$2/2</f></c>
      <c r="H1"><f>ROUND(Other!B2/3,2)</f></c>
      <c r="I1"><f>ROUNDUP(Other!B2/3,2)</f></c>
      <c r="J1"><f>ROUNDDOWN(Other!B2/3,2)</f></c>
      <c r="K1"><f>INT(Other!B2/3)</f></c>
    </row>
  </sheetData>
</worksheet>`;
    const sheet2Xml = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="2">
      <c r="B2"><v>8</v></c>
    </row>
  </sheetData>
</worksheet>`;
    const zip = createStoredZip([
      { name: "[Content_Types].xml", data: `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>` },
      { name: "xl/workbook.xml", data: workbookXml },
      { name: "xl/_rels/workbook.xml.rels", data: workbookRelsXml },
      { name: "xl/worksheets/sheet1.xml", data: sheet1Xml },
      { name: "xl/worksheets/sheet2.xml", data: sheet2Xml }
    ]);

    const workbook = await api.parseWorkbook(zip, "arith.xlsx");
    const summarySheet = workbook.sheets[0];
    const c1 = summarySheet.cells.find((cell) => cell.address === "C1");
    const d1 = summarySheet.cells.find((cell) => cell.address === "D1");
    const e1 = summarySheet.cells.find((cell) => cell.address === "E1");
    const f1 = summarySheet.cells.find((cell) => cell.address === "F1");
    const g1 = summarySheet.cells.find((cell) => cell.address === "G1");
    const h1 = summarySheet.cells.find((cell) => cell.address === "H1");
    const i1 = summarySheet.cells.find((cell) => cell.address === "I1");
    const j1 = summarySheet.cells.find((cell) => cell.address === "J1");
    const k1 = summarySheet.cells.find((cell) => cell.address === "K1");

    expect(c1?.outputValue).toBe("15");
    expect(d1?.outputValue).toBe("10");
    expect(e1?.outputValue).toBe("4");
    expect(f1?.outputValue).toBe("15");
    expect(g1?.outputValue).toBe("4");
    expect(h1?.outputValue).toBe("2.67");
    expect(i1?.outputValue).toBe("2.67");
    expect(j1?.outputValue).toBe("2.66");
    expect(k1?.outputValue).toBe("2");
    expect(c1?.resolutionStatus).toBe("resolved");
    expect(d1?.resolutionStatus).toBe("resolved");
    expect(e1?.resolutionStatus).toBe("resolved");
    expect(f1?.resolutionStatus).toBe("resolved");
    expect(g1?.resolutionStatus).toBe("resolved");
    expect(h1?.resolutionStatus).toBe("resolved");
    expect(i1?.resolutionStatus).toBe("resolved");
    expect(j1?.resolutionStatus).toBe("resolved");
    expect(k1?.resolutionStatus).toBe("resolved");
  });

  it("expands shared formulas and resolves follower cells", async () => {
    const api = bootCore();
    const workbookXml = `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Summary" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`;
    const workbookRelsXml = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`;
    const sheetXml = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><v>10</v></c>
      <c r="B1"><f t="shared" ref="B1:B3" si="0">A1+1</f><v>11</v></c>
    </row>
    <row r="2">
      <c r="A2"><v>20</v></c>
      <c r="B2"><f t="shared" si="0"/><v>21</v></c>
    </row>
    <row r="3">
      <c r="A3"><v>30</v></c>
      <c r="B3"><f t="shared" si="0"/></c>
    </row>
  </sheetData>
</worksheet>`;
    const zip = createStoredZip([
      { name: "[Content_Types].xml", data: `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>` },
      { name: "xl/workbook.xml", data: workbookXml },
      { name: "xl/_rels/workbook.xml.rels", data: workbookRelsXml },
      { name: "xl/worksheets/sheet1.xml", data: sheetXml }
    ]);

    const workbook = await api.parseWorkbook(zip, "shared-formula.xlsx");
    const summarySheet = workbook.sheets[0];

    expect(summarySheet.cells.find((cell) => cell.address === "B1")?.formulaText).toBe("=A1+1");
    expect(summarySheet.cells.find((cell) => cell.address === "B2")?.formulaText).toBe("=A2+1");
    expect(summarySheet.cells.find((cell) => cell.address === "B3")?.formulaText).toBe("=A3+1");
    expect(summarySheet.cells.find((cell) => cell.address === "B3")?.outputValue).toBe("31");
    expect(summarySheet.cells.find((cell) => cell.address === "B3")?.resolutionStatus).toBe("resolved");
  });

  it("evaluates SUM AVERAGE MIN MAX over ranges", async () => {
    const api = bootCore();
    const workbookXml = `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Summary" sheetId="1" r:id="rId1"/>
    <sheet name="Other" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>`;
    const workbookRelsXml = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>`;
    const sheet1Xml = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><v>1</v></c>
      <c r="B1"><v>2</v></c>
      <c r="C1" t="inlineStr"><is><t>text</t></is></c>
    </row>
    <row r="2">
      <c r="A2"><v>3</v></c>
      <c r="B2"><v>4</v></c>
    </row>
    <row r="4">
      <c r="A4"><f>SUM(A1:B2)</f></c>
      <c r="B4"><f>AVERAGE(A1:B2)</f></c>
      <c r="C4"><f>MIN(A1:B2)</f></c>
      <c r="D4"><f>MAX(Other!A1:B2)</f></c>
      <c r="E4"><f>COUNT(A1:C2)</f></c>
      <c r="F4"><f>COUNTA(A1:C2)</f></c>
      <c r="G4"><f>SUM(A1,B1,10)</f></c>
      <c r="H4"><f>SUM($A$1:$B$2)</f></c>
      <c r="I4"><f>VLOOKUP(2,Other!A1:B2,2,FALSE)</f></c>
      <c r="J4"><f>VLOOKUP(3,$A$1:$B$2,2,0)</f></c>
      <c r="K4"><f>COUNTIF(A1:B2,">2")</f></c>
      <c r="L4"><f>SUMIF(A1:B2,">2")</f></c>
      <c r="M4"><f>COUNTIF(A1:C2,"text")</f></c>
      <c r="N4"><f>AVERAGEIF(A1:B2,">2")</f></c>
      <c r="O4"><f>COUNTIFS(A1:B2,"&gt;1",A1:B2,"&lt;4")</f></c>
      <c r="P4"><f>SUMIFS(A1:B2,A1:B2,"&gt;1",A1:B2,"&lt;4")</f></c>
      <c r="Q4"><f>AVERAGEIFS(A1:B2,A1:B2,"&gt;1",A1:B2,"&lt;4")</f></c>
      <c r="R4"><f>HLOOKUP("K2",A6:C7,2,FALSE)</f></c>
      <c r="S4"><f>MATCH("K2",A6:C6,0)</f></c>
      <c r="T4"><f>INDEX(A7:C7,1,2)</f></c>
      <c r="U4"><f>INDEX(A7:C7,1,MATCH("K2",A6:C6,0))</f></c>
      <c r="V4"><f>XLOOKUP("K2",A6:C6,A7:C7)</f></c>
      <c r="W4"><f>XLOOKUP("ZZ",A6:C6,A7:C7,"NF")</f></c>
      <c r="X4"><f>CHOOSE(2,"A","B","C")</f></c>
      <c r="Y4"><f>CHOOSE(MATCH("K2",A6:C6,0),"A","B","C")</f></c>
    </row>
    <row r="6">
      <c r="A6" t="inlineStr"><is><t>K1</t></is></c>
      <c r="B6" t="inlineStr"><is><t>K2</t></is></c>
      <c r="C6" t="inlineStr"><is><t>K3</t></is></c>
    </row>
    <row r="7">
      <c r="A7"><v>11</v></c>
      <c r="B7"><v>22</v></c>
      <c r="C7"><v>33</v></c>
    </row>
  </sheetData>
</worksheet>`;
    const sheet2Xml = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><v>1</v></c>
      <c r="B1"><v>7</v></c>
    </row>
    <row r="2">
      <c r="A2"><v>2</v></c>
      <c r="B2"><v>8</v></c>
    </row>
  </sheetData>
</worksheet>`;
    const zip = createStoredZip([
      { name: "[Content_Types].xml", data: `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>` },
      { name: "xl/workbook.xml", data: workbookXml },
      { name: "xl/_rels/workbook.xml.rels", data: workbookRelsXml },
      { name: "xl/worksheets/sheet1.xml", data: sheet1Xml },
      { name: "xl/worksheets/sheet2.xml", data: sheet2Xml }
    ]);

    const workbook = await api.parseWorkbook(zip, "func.xlsx");
    const summarySheet = workbook.sheets[0];

    expect(summarySheet.cells.find((cell) => cell.address === "A4")?.outputValue).toBe("10");
    expect(summarySheet.cells.find((cell) => cell.address === "B4")?.outputValue).toBe("2.5");
    expect(summarySheet.cells.find((cell) => cell.address === "C4")?.outputValue).toBe("1");
    expect(summarySheet.cells.find((cell) => cell.address === "D4")?.outputValue).toBe("8");
    expect(summarySheet.cells.find((cell) => cell.address === "E4")?.outputValue).toBe("4");
    expect(summarySheet.cells.find((cell) => cell.address === "F4")?.outputValue).toBe("5");
    expect(summarySheet.cells.find((cell) => cell.address === "G4")?.outputValue).toBe("13");
    expect(summarySheet.cells.find((cell) => cell.address === "H4")?.outputValue).toBe("10");
    expect(summarySheet.cells.find((cell) => cell.address === "I4")?.outputValue).toBe("8");
    expect(summarySheet.cells.find((cell) => cell.address === "J4")?.outputValue).toBe("4");
    expect(summarySheet.cells.find((cell) => cell.address === "K4")?.outputValue).toBe("2");
    expect(summarySheet.cells.find((cell) => cell.address === "L4")?.outputValue).toBe("7");
    expect(summarySheet.cells.find((cell) => cell.address === "M4")?.outputValue).toBe("1");
    expect(summarySheet.cells.find((cell) => cell.address === "N4")?.outputValue).toBe("3.5");
    expect(summarySheet.cells.find((cell) => cell.address === "O4")?.outputValue).toBe("2");
    expect(summarySheet.cells.find((cell) => cell.address === "P4")?.outputValue).toBe("5");
    expect(summarySheet.cells.find((cell) => cell.address === "Q4")?.outputValue).toBe("2.5");
    expect(summarySheet.cells.find((cell) => cell.address === "R4")?.outputValue).toBe("22");
    expect(summarySheet.cells.find((cell) => cell.address === "S4")?.outputValue).toBe("2");
    expect(summarySheet.cells.find((cell) => cell.address === "T4")?.outputValue).toBe("22");
    expect(summarySheet.cells.find((cell) => cell.address === "U4")?.outputValue).toBe("22");
    expect(summarySheet.cells.find((cell) => cell.address === "V4")?.outputValue).toBe("22");
    expect(summarySheet.cells.find((cell) => cell.address === "W4")?.outputValue).toBe("NF");
    expect(summarySheet.cells.find((cell) => cell.address === "X4")?.outputValue).toBe("B");
    expect(summarySheet.cells.find((cell) => cell.address === "Y4")?.outputValue).toBe("B");
  });

  it("evaluates comparison expressions and IF formulas", async () => {
    const api = bootCore();
    const workbookXml = `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Summary" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`;
    const workbookRelsXml = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`;
    const sheetXml = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><v>10</v></c>
      <c r="B1"><v>5</v></c>
      <c r="C1"><f>A1&gt;B1</f></c>
      <c r="D1"><f>A1=B1</f></c>
      <c r="E1"><f>IF(A1&gt;B1,1,0)</f></c>
      <c r="F1"><f>IF(B1&gt;A1,"OK","NG")</f></c>
      <c r="G1"><f>IF((A1-B1)=5,"MATCH","MISS")</f></c>
      <c r="H1"><f>AND(A1&gt;B1,B1&gt;0)</f></c>
      <c r="I1"><f>OR(A1&lt;B1,B1&gt;0)</f></c>
      <c r="J1"><f>NOT(A1=B1)</f></c>
      <c r="K1"><f>IF(AND(A1&gt;B1,B1&gt;0),"OK","NG")</f></c>
      <c r="L1"><f>"ID-"&amp;A1</f></c>
      <c r="M1"><f>A1&amp;" / "&amp;B1</f></c>
      <c r="N1"><f>"Result:"&amp;IF(A1&gt;B1,"OK","NG")</f></c>
      <c r="O1"><f>"ABS-"&amp;$A$1&amp;"-"&amp;$B1</f></c>
      <c r="P1"><f>LEFT("ABCDE",2)</f></c>
      <c r="Q1"><f>RIGHT("ABCDE",3)</f></c>
      <c r="R1"><f>MID("ABCDE",2,2)</f></c>
      <c r="S1"><f>LEN("AB CD")</f></c>
      <c r="T1"><f>TRIM("  A   B  ")</f></c>
      <c r="U1"><f>LEFT("ID-"&amp;A1,4)</f></c>
      <c r="V1"><f>SUBSTITUTE("A-B-B","B","X")</f></c>
      <c r="W1"><f>SUBSTITUTE("A-B-B","B","X",2)</f></c>
      <c r="X1"><f>REPLACE("ABCDE",2,2,"ZZ")</f></c>
      <c r="Y1"><f>TEXT(1234.5,"#,##0.00")</f></c>
      <c r="Z1"><f>TEXT(45292,"yyyy-mm-dd")</f></c>
      <c r="AA1"><f>TEXT(0.5,"hh:mm:ss")</f></c>
      <c r="AB1"><f>YEAR(45292)</f></c>
      <c r="AC1"><f>MONTH(45292)</f></c>
      <c r="AD1"><f>DAY(45292)</f></c>
      <c r="AE1"><f>YEAR(TEXT(45292,"yyyy-mm-dd"))</f></c>
      <c r="AF1"><f>DATE(2024,1,1)</f></c>
      <c r="AG1"><f>TEXT(DATE(2024,1,1),"yyyy-mm-dd")</f></c>
      <c r="AH1"><f>VALUE("1,234.50")</f></c>
      <c r="AI1"><f>VALUE("2024-01-01")</f></c>
      <c r="AJ1"><f>IFERROR(Unknown!A1,"ALT")</f></c>
      <c r="AK1"><f>IFERROR(A1+B1,"ALT")</f></c>
      <c r="AL1"><f>ISBLANK(C2)</f></c>
      <c r="AM1"><f>ISNUMBER(A1)</f></c>
      <c r="AN1"><f>ISTEXT("HELLO")</f></c>
      <c r="AO1"><f>IF(ISBLANK(C2),"EMPTY","FILLED")</f></c>
      <c r="AP1"><f>ISERROR(Unknown!A1)</f></c>
      <c r="AQ1"><f>ISNA(VLOOKUP(99,A1:B1,2,FALSE))</f></c>
    </row>
  </sheetData>
</worksheet>`;
    const zip = createStoredZip([
      { name: "[Content_Types].xml", data: `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>` },
      { name: "xl/workbook.xml", data: workbookXml },
      { name: "xl/_rels/workbook.xml.rels", data: workbookRelsXml },
      { name: "xl/worksheets/sheet1.xml", data: sheetXml }
    ]);

    const workbook = await api.parseWorkbook(zip, "if.xlsx");
    const summarySheet = workbook.sheets[0];

    expect(summarySheet.cells.find((cell) => cell.address === "C1")?.outputValue).toBe("TRUE");
    expect(summarySheet.cells.find((cell) => cell.address === "D1")?.outputValue).toBe("FALSE");
    expect(summarySheet.cells.find((cell) => cell.address === "E1")?.outputValue).toBe("1");
    expect(summarySheet.cells.find((cell) => cell.address === "F1")?.outputValue).toBe("NG");
    expect(summarySheet.cells.find((cell) => cell.address === "G1")?.outputValue).toBe("MATCH");
    expect(summarySheet.cells.find((cell) => cell.address === "H1")?.outputValue).toBe("TRUE");
    expect(summarySheet.cells.find((cell) => cell.address === "I1")?.outputValue).toBe("TRUE");
    expect(summarySheet.cells.find((cell) => cell.address === "J1")?.outputValue).toBe("TRUE");
    expect(summarySheet.cells.find((cell) => cell.address === "K1")?.outputValue).toBe("OK");
    expect(summarySheet.cells.find((cell) => cell.address === "L1")?.outputValue).toBe("ID-10");
    expect(summarySheet.cells.find((cell) => cell.address === "M1")?.outputValue).toBe("10 / 5");
    expect(summarySheet.cells.find((cell) => cell.address === "N1")?.outputValue).toBe("Result:OK");
    expect(summarySheet.cells.find((cell) => cell.address === "O1")?.outputValue).toBe("ABS-10-5");
    expect(summarySheet.cells.find((cell) => cell.address === "P1")?.outputValue).toBe("AB");
    expect(summarySheet.cells.find((cell) => cell.address === "Q1")?.outputValue).toBe("CDE");
    expect(summarySheet.cells.find((cell) => cell.address === "R1")?.outputValue).toBe("BC");
    expect(summarySheet.cells.find((cell) => cell.address === "S1")?.outputValue).toBe("5");
    expect(summarySheet.cells.find((cell) => cell.address === "T1")?.outputValue).toBe("A B");
    expect(summarySheet.cells.find((cell) => cell.address === "U1")?.outputValue).toBe("ID-1");
    expect(summarySheet.cells.find((cell) => cell.address === "V1")?.outputValue).toBe("A-X-X");
    expect(summarySheet.cells.find((cell) => cell.address === "W1")?.outputValue).toBe("A-B-X");
    expect(summarySheet.cells.find((cell) => cell.address === "X1")?.outputValue).toBe("AZZDE");
    expect(summarySheet.cells.find((cell) => cell.address === "Y1")?.outputValue).toBe("1,234.50");
    expect(summarySheet.cells.find((cell) => cell.address === "Z1")?.outputValue).toBe("2024-01-01");
    expect(summarySheet.cells.find((cell) => cell.address === "AA1")?.outputValue).toBe("12:00:00");
    expect(summarySheet.cells.find((cell) => cell.address === "AB1")?.outputValue).toBe("2024");
    expect(summarySheet.cells.find((cell) => cell.address === "AC1")?.outputValue).toBe("1");
    expect(summarySheet.cells.find((cell) => cell.address === "AD1")?.outputValue).toBe("1");
    expect(summarySheet.cells.find((cell) => cell.address === "AE1")?.outputValue).toBe("2024");
    expect(summarySheet.cells.find((cell) => cell.address === "AF1")?.outputValue).toBe("45292");
    expect(summarySheet.cells.find((cell) => cell.address === "AG1")?.outputValue).toBe("2024-01-01");
    expect(summarySheet.cells.find((cell) => cell.address === "AH1")?.outputValue).toBe("1234.5");
    expect(summarySheet.cells.find((cell) => cell.address === "AI1")?.outputValue).toBe("45292");
    expect(summarySheet.cells.find((cell) => cell.address === "AJ1")?.outputValue).toBe("ALT");
    expect(summarySheet.cells.find((cell) => cell.address === "AK1")?.outputValue).toBe("15");
    expect(summarySheet.cells.find((cell) => cell.address === "AL1")?.outputValue).toBe("TRUE");
    expect(summarySheet.cells.find((cell) => cell.address === "AM1")?.outputValue).toBe("TRUE");
    expect(summarySheet.cells.find((cell) => cell.address === "AN1")?.outputValue).toBe("TRUE");
    expect(summarySheet.cells.find((cell) => cell.address === "AO1")?.outputValue).toBe("EMPTY");
    expect(summarySheet.cells.find((cell) => cell.address === "AP1")?.outputValue).toBe("TRUE");
    expect(summarySheet.cells.find((cell) => cell.address === "AQ1")?.outputValue).toBe("TRUE");
  });

  it("renders accounting zero format as dash instead of currency zero", () => {
    const api = bootCore();
    const workbook = { name: "accounting.xlsx" };
    const sheet = {
      name: "Summary",
      index: 1,
      path: "xl/worksheets/sheet1.xml",
      merges: [],
      tables: [],
      images: [],
      maxRow: 2,
      maxCol: 2,
      cells: [
        { row: 1, col: 1, address: "A1", valueType: "str", rawValue: "費用", outputValue: "費用", formulaText: "", resolutionStatus: null, styleIndex: 0, borders: { top: true, bottom: true, left: true, right: true }, numFmtId: 0, formatCode: "General" },
        { row: 1, col: 2, address: "B1", valueType: "str", rawValue: "合計費用", outputValue: "合計費用", formulaText: "", resolutionStatus: null, styleIndex: 0, borders: { top: true, bottom: true, left: true, right: true }, numFmtId: 0, formatCode: "General" },
        { row: 2, col: 1, address: "A2", valueType: "n", rawValue: "0", outputValue: "¥ -", formulaText: "", resolutionStatus: null, styleIndex: 0, borders: { top: true, bottom: true, left: true, right: true }, numFmtId: 44, formatCode: "_ \"¥\"* #,##0.00_ ;_ \"¥\"* \\-#,##0.00_ ;_ \"¥\"* \"-\"??_ ;_ @_ " },
        { row: 2, col: 2, address: "B2", valueType: "n", rawValue: "0", outputValue: "¥ -", formulaText: "", resolutionStatus: null, styleIndex: 0, borders: { top: true, bottom: true, left: true, right: true }, numFmtId: 44, formatCode: "_ \"¥\"* #,##0.00_ ;_ \"¥\"* \\-#,##0.00_ ;_ \"¥\"* \"-\"??_ ;_ @_ " }
      ]
    };

    const markdownFile = api.convertSheetToMarkdown(workbook, sheet, {
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true,
      treatFirstRowAsHeader: true
    });

    expect(markdownFile.markdown).toContain("| ¥ - | ¥ - |");
  });

  it("includes formula resolution source in diagnostics", () => {
    const api = bootCore();
    const workbook = { name: "source.xlsx" };
    const sheet = {
      name: "Formula Source",
      index: 1,
      path: "xl/worksheets/sheet1.xml",
      merges: [],
      tables: [],
      images: [],
      maxRow: 2,
      maxCol: 2,
      cells: [
        { row: 1, col: 1, address: "A1", valueType: "str", rawValue: "label", outputValue: "label", formulaText: "", resolutionStatus: null, resolutionSource: null, cachedValueState: null, styleIndex: 0, borders: { top: true, bottom: true, left: true, right: true }, numFmtId: 0, formatCode: "General", spillRef: "" },
        { row: 1, col: 2, address: "B1", valueType: "formula", rawValue: "2", outputValue: "2", formulaText: "=1+1", resolutionStatus: "resolved", resolutionSource: "ast_evaluator", cachedValueState: "absent", styleIndex: 0, borders: { top: true, bottom: true, left: true, right: true }, numFmtId: 0, formatCode: "General", spillRef: "" },
        { row: 2, col: 1, address: "A2", valueType: "formula", rawValue: "=UNKNOWN(A1)", outputValue: "=UNKNOWN(A1)", formulaText: "=UNKNOWN(A1)", resolutionStatus: "fallback_formula", resolutionSource: "formula_text", cachedValueState: "absent", styleIndex: 0, borders: { top: true, bottom: true, left: true, right: true }, numFmtId: 0, formatCode: "General", spillRef: "" },
        { row: 2, col: 2, address: "B2", valueType: "formula", rawValue: "#REF!", outputValue: "=[other.xlsx]Sheet1!A1", formulaText: "=[other.xlsx]Sheet1!A1", resolutionStatus: "unsupported_external", resolutionSource: "external_unsupported", cachedValueState: "absent", styleIndex: 0, borders: { top: true, bottom: true, left: true, right: true }, numFmtId: 0, formatCode: "General", spillRef: "" }
      ]
    };

    const markdownFile = api.convertSheetToMarkdown(workbook, sheet, {
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true,
      treatFirstRowAsHeader: true
    });

    expect(markdownFile.summary.formulaDiagnostics).toEqual([
      { address: "B1", formulaText: "=1+1", status: "resolved", source: "ast_evaluator", outputValue: "2" },
      { address: "A2", formulaText: "=UNKNOWN(A1)", status: "fallback_formula", source: "formula_text", outputValue: "=UNKNOWN(A1)" },
      { address: "B2", formulaText: "=[other.xlsx]Sheet1!A1", status: "unsupported_external", source: "external_unsupported", outputValue: "=[other.xlsx]Sheet1!A1" }
    ]);
  });

  it("does not treat wide sparse merge-heavy layout regions as tables", () => {
    const api = bootCore();
    const workbook = { name: "layout.xlsx" };
    const cells = [];
    for (let row = 1; row <= 5; row += 1) {
      for (let col = 1; col <= 12; col += 1) {
        const address = `${String.fromCharCode(64 + col)}${row}`;
        let outputValue = "";
        if (col === 1) outputValue = `label${row}`;
        if (row === 4 && col === 8) outputValue = "終了日時";
        cells.push({
          row,
          col,
          address,
          valueType: outputValue ? "str" : "n",
          rawValue: outputValue,
          outputValue,
          formulaText: "",
          resolutionStatus: null,
          resolutionSource: null,
          styleIndex: 0,
          borders: { top: true, bottom: true, left: true, right: true },
          numFmtId: 0,
          formatCode: "General",
          formulaType: "",
          spillRef: ""
        });
      }
    }
    const sheet = {
      name: "Layout",
      index: 1,
      path: "xl/worksheets/sheet1.xml",
      merges: [
        { startRow: 1, startCol: 2, endRow: 1, endCol: 12, ref: "B1:L1" },
        { startRow: 2, startCol: 2, endRow: 2, endCol: 12, ref: "B2:L2" },
        { startRow: 3, startCol: 2, endRow: 3, endCol: 12, ref: "B3:L3" },
        { startRow: 4, startCol: 2, endRow: 4, endCol: 7, ref: "B4:G4" },
        { startRow: 4, startCol: 8, endRow: 4, endCol: 12, ref: "H4:L4" },
        { startRow: 5, startCol: 2, endRow: 5, endCol: 12, ref: "B5:L5" }
      ],
      tables: [],
      images: [],
      maxRow: 5,
      maxCol: 12,
      cells
    };

    const markdownFile = api.convertSheetToMarkdown(workbook, sheet, {
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true,
      treatFirstRowAsHeader: true
    });

    expect(markdownFile.summary.tableScores).toHaveLength(0);
    expect(markdownFile.markdown).not.toContain("### Table 001");
    expect(markdownFile.markdown).toContain("label1");
    expect(markdownFile.markdown).toContain("終了日時");
  });

  it("splits distant same-row narrative cells into separate blocks", () => {
    const api = bootCore();
    const workbook = { name: "narrative-gap.xlsx" };
    const sheet = {
      name: "Layout",
      index: 1,
      path: "xl/worksheets/sheet1.xml",
      merges: [],
      tables: [],
      images: [],
      maxRow: 1,
      maxCol: 10,
      cells: [
        {
          row: 1, col: 1, address: "A1", valueType: "str", rawValue: "イベント チェックリスト", outputValue: "イベント チェックリスト",
          formulaText: "", resolutionStatus: null, resolutionSource: null, styleIndex: 0,
          borders: { top: false, bottom: false, left: false, right: false }, numFmtId: 0, formatCode: "General", formulaType: "", spillRef: ""
        },
        {
          row: 1, col: 8, address: "H1", valueType: "str", rawValue: "イベント カテゴリ", outputValue: "イベント カテゴリ",
          formulaText: "", resolutionStatus: null, resolutionSource: null, styleIndex: 0,
          borders: { top: false, bottom: false, left: false, right: false }, numFmtId: 0, formatCode: "General", formulaType: "", spillRef: ""
        }
      ]
    };

    const markdownFile = api.convertSheetToMarkdown(workbook, sheet, {
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true,
      treatFirstRowAsHeader: true
    });

    expect(markdownFile.summary.narrativeBlocks).toBe(2);
    expect(markdownFile.markdown).toContain("イベント チェックリスト");
    expect(markdownFile.markdown).toContain("イベント カテゴリ");
    expect(markdownFile.markdown).not.toContain("イベント チェックリスト イベント カテゴリ");
  });

  it("treats empty-string cached formula results as cached", async () => {
    const api = bootCore();
    const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`;
    const workbookRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`;
    const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><f>IF(1=1,"","X")</f><v></v></c>
    </row>
  </sheetData>
</worksheet>`;

    const zip = createStoredZip([
      { name: "[Content_Types].xml", data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>` },
      { name: "xl/workbook.xml", data: workbookXml },
      { name: "xl/_rels/workbook.xml.rels", data: workbookRelsXml },
      { name: "xl/worksheets/sheet1.xml", data: sheetXml }
    ]);

    const workbook = await api.parseWorkbook(zip, "empty-cache.xlsx");
    const cell = workbook.sheets[0].cells.find((entry) => entry.address === "A1");
    const markdownFile = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true
    })[0];

    expect(cell?.rawValue).toBe("");
    expect(cell?.outputValue).toBe("");
    expect(cell?.resolutionStatus).toBe("resolved");
    expect(cell?.resolutionSource).toBe("cached_value");
    expect(cell?.cachedValueState).toBe("present_empty");
    expect(markdownFile.summary.formulaDiagnostics).toEqual([
      { address: "A1", formulaText: "=IF(1=1,\"\",\"X\")", status: "resolved", source: "cached_value", outputValue: "" }
    ]);
  });

  it("renders external and workbook hyperlinks into markdown", async () => {
    const api = bootCore();
    const zip = createStoredZip([
      {
        name: "[Content_Types].xml",
        data: `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`
      },
      {
        name: "_rels/.rels",
        data: `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`
      },
      {
        name: "xl/workbook.xml",
        data: `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Summary" sheetId="1" r:id="rId1"/>
    <sheet name="Other" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>`
      },
      {
        name: "xl/_rels/workbook.xml.rels",
        data: `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>`
      },
      {
        name: "xl/worksheets/sheet1.xml",
        data: `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Open</t></is></c>
    </row>
    <row r="2">
      <c r="A2" t="inlineStr"><is><t>Jump</t></is></c>
    </row>
  </sheetData>
  <hyperlinks>
    <hyperlink ref="A1" r:id="rId1"/>
    <hyperlink ref="A2" location="Other!A1"/>
  </hyperlinks>
</worksheet>`
      },
      {
        name: "xl/worksheets/_rels/sheet1.xml.rels",
        data: `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/" TargetMode="External"/>
</Relationships>`
      },
      {
        name: "xl/worksheets/sheet2.xml",
        data: `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Target</t></is></c>
    </row>
  </sheetData>
</worksheet>`
      }
    ]);

    const workbook = await api.parseWorkbook(zip, "link-book.xlsx");
    const markdownFiles = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true,
      formattingMode: "github"
    });

    expect(markdownFiles[0].markdown).toContain("[Open](https://example.com/)");
    expect(markdownFiles[0].markdown).toContain("[Jump](#link-book_002_Other_github) (Other!A1)");
    expect(markdownFiles[0].markdown).not.toContain("<ins>Open</ins>");
    expect(markdownFiles[0].markdown).not.toContain("<ins>Jump</ins>");
    expect(markdownFiles[1].markdown).toContain('<a id="link-book_002_Other_github"></a>');
  });

  it("preserves supported rich text as github-compatible markdown", async () => {
    const api = bootCore();
    const workbookBuffer = createStoredZip([
      {
        name: "[Content_Types].xml",
        data: `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`
      },
      {
        name: "_rels/.rels",
        data: `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`
      },
      {
        name: "xl/workbook.xml",
        data: `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`
      },
      {
        name: "xl/_rels/workbook.xml.rels",
        data: `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`
      },
      {
        name: "xl/worksheets/sheet1.xml",
        data: `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1" t="inlineStr" s="1"><is><t>WholeCell</t></is></c>
    </row>
  </sheetData>
</worksheet>`
      },
      {
        name: "xl/sharedStrings.xml",
        data: `<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <r><rPr><b/></rPr><t>Bold</t></r>
    <r><rPr><i/></rPr><t>Italic</t></r>
    <r><rPr><strike/></rPr><t>Strike</t></r>
    <r><rPr><u/></rPr><t>Under</t></r>
  </si>
</sst>`
      },
      {
        name: "xl/styles.xml",
        data: `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font/>
    <font><u/></font>
  </fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
  </cellXfs>
</styleSheet>`
      }
    ]);

    const workbook = await api.parseWorkbook(workbookBuffer, "rich-text.xlsx");
    const githubFiles = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true,
      formattingMode: "github"
    });
    const plainFiles = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true,
      formattingMode: "plain"
    });

    expect(githubFiles[0].markdown).toContain("**Bold***Italic*~~Strike~~<ins>Under</ins>");
    expect(githubFiles[0].markdown).toContain("<ins>WholeCell</ins>");
    expect(githubFiles[0].summary.formattingMode).toBe("github");
    expect(plainFiles[0].markdown).toContain("BoldItalicStrikeUnder");
    expect(plainFiles[0].markdown).not.toContain("<ins>WholeCell</ins>");
  });

  it("converts the rich text fixture into github-compatible markdown", async () => {
    const api = bootCore();
    const fixturePath = path.resolve(__dirname, "./fixtures/rich/rich-text-github-sample01.xlsx");
    const fixtureBytes = readFileSync(fixturePath);
    const arrayBuffer = fixtureBytes.buffer.slice(
      fixtureBytes.byteOffset,
      fixtureBytes.byteOffset + fixtureBytes.byteLength
    );

    const workbook = await api.parseWorkbook(arrayBuffer, "rich-text-github-sample01.xlsx");
    const markdownFile = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true,
      formattingMode: "github"
    })[0];

    expect(markdownFile.summary.formattingMode).toBe("github");
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
});
