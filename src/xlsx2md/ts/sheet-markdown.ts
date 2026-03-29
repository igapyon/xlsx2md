/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */

(() => {
  const moduleRegistry = getXlsx2mdModuleRegistry();
  type FormulaResolutionStatus = "resolved" | "fallback_formula" | "unsupported_external" | null;

  type BorderFlags = {
    top: boolean;
    bottom: boolean;
    left: boolean;
    right: boolean;
  };

  type ParsedCell = {
    address: string;
    row: number;
    col: number;
    valueType: string;
    rawValue: string;
    outputValue: string;
    formulaText: string;
    resolutionStatus: FormulaResolutionStatus;
    resolutionSource: string | null;
    styleIndex: number;
    borders: BorderFlags;
    numFmtId: number;
    formatCode: string;
    textStyle: {
      bold: boolean;
      italic: boolean;
      strike: boolean;
      underline: boolean;
    };
    richTextRuns: Array<{
      text: string;
      bold: boolean;
      italic: boolean;
      strike: boolean;
      underline: boolean;
    }> | null;
    formulaType: string;
    spillRef: string;
    hyperlink: {
      kind: "external" | "internal";
      target: string;
      location: string;
      tooltip: string;
      display: string;
    } | null;
  };

  type MergeRange = {
    startRow: number;
    startCol: number;
    endRow: number;
    endCol: number;
    ref: string;
  };

  type NarrativeItem = {
    row: number;
    startCol: number;
    text: string;
    cellValues: string[];
  };

  type NarrativeBlock = {
    startRow: number;
    startCol: number;
    endRow: number;
    lines: string[];
    items: NarrativeItem[];
  };

  type SectionBlock = {
    startRow: number;
    startCol: number;
    endRow: number;
    endCol: number;
  };

  type ParsedImageAsset = {
    anchor: string;
    filename: string;
    path: string;
    data: Uint8Array;
  };

  type ParsedChartAsset = {
    anchor: string;
    title: string;
    chartType: string;
    series: {
      name: string;
      categoriesRef: string;
      valuesRef: string;
      axis: "primary" | "secondary";
    }[];
  };

  type ParsedShapeAsset = {
    anchor: string;
    rawEntries: {
      key: string;
      value: string;
    }[];
    svgFilename: string | null;
    svgPath: string | null;
    svgData: Uint8Array | null;
  };

  type ParsedSheet = {
    name: string;
    index: number;
    cells: ParsedCell[];
    merges: MergeRange[];
    images: ParsedImageAsset[];
    charts: ParsedChartAsset[];
    shapes: ParsedShapeAsset[];
  };

  type ParsedWorkbook = {
    name: string;
    sheets: ParsedSheet[];
  };

  type MarkdownOptions = {
    treatFirstRowAsHeader?: boolean;
    trimText?: boolean;
    removeEmptyRows?: boolean;
    removeEmptyColumns?: boolean;
    includeShapeDetails?: boolean;
    outputMode?: "display" | "raw" | "both";
    formattingMode?: "plain" | "github";
    tableDetectionMode?: "balanced" | "border";
  };

  type TableCandidate = {
    startRow: number;
    startCol: number;
    endRow: number;
    endCol: number;
    score: number;
    reasonSummary: string[];
  };

  type ShapeBlock = {
    startRow: number;
    startCol: number;
    endRow: number;
    endCol: number;
    shapeIndexes: number[];
  };

  type MarkdownFile = {
    fileName: string;
    sheetName: string;
    markdown: string;
    summary: {
      outputMode: "display" | "raw" | "both";
      formattingMode: "plain" | "github";
      tableDetectionMode: "balanced" | "border";
      sections: number;
      tables: number;
      narrativeBlocks: number;
      merges: number;
      images: number;
      charts: number;
      cells: number;
      tableScores: Array<{
        range: string;
        score: number;
        reasons: string[];
      }>;
      formulaDiagnostics: Array<{
        address: string;
        formulaText: string;
        status: FormulaResolutionStatus;
        source: string | null;
        outputValue: string;
      }>;
    };
  };

  type SheetMarkdownDeps = {
    renderNarrativeBlock: (block: NarrativeBlock) => string;
    detectTableCandidates: (
      sheet: ParsedSheet,
      buildCellMap: (sheet: ParsedSheet) => Map<string, ParsedCell>,
      tableDetectionMode?: "balanced" | "border"
    ) => TableCandidate[];
    matrixFromCandidate: (
      sheet: ParsedSheet,
      candidate: TableCandidate,
      options: MarkdownOptions,
      buildCellMap: (sheet: ParsedSheet) => Map<string, ParsedCell>,
      formatCellForMarkdown: (cell: ParsedCell | undefined, options: MarkdownOptions) => string
    ) => string[][];
    renderMarkdownTable: (rows: string[][], treatFirstRowAsHeader: boolean) => string;
    createOutputFileName: (
      workbookName: string,
      sheetIndex: number,
      sheetName: string,
      outputMode?: "display" | "raw" | "both",
      formattingMode?: "plain" | "github"
    ) => string;
    extractShapeBlocks: (
      shapes: ParsedShapeAsset[],
      options: {
        defaultCellWidthEmu: number;
        defaultCellHeightEmu: number;
        shapeBlockGapXEmu: number;
        shapeBlockGapYEmu: number;
      }
    ) => ShapeBlock[];
    renderHierarchicalRawEntries: (entries: { key: string; value: string }[]) => string[];
    parseCellAddress: (address: string) => { row: number; col: number };
    formatRange: (startRow: number, startCol: number, endRow: number, endCol: number) => string;
    colToLetters: (col: number) => string;
    normalizeMarkdownText?: (text: string) => string;
    defaultCellWidthEmu: number;
    defaultCellHeightEmu: number;
    shapeBlockGapXEmu: number;
    shapeBlockGapYEmu: number;
  };

  function createSheetMarkdownApi(deps: SheetMarkdownDeps) {
    const richTextRenderer = requireXlsx2mdRichTextRendererModule<ParsedCell>().createRichTextRendererApi({
      normalizeMarkdownText: deps.normalizeMarkdownText
    });

    function buildCellMap(sheet: ParsedSheet): Map<string, ParsedCell> {
      const map = new Map<string, ParsedCell>();
      for (const cell of sheet.cells) {
        map.set(`${cell.row}:${cell.col}`, cell);
      }
      return map;
    }

    function createSheetAnchorId(
      workbookName: string,
      sheetIndex: number,
      sheetName: string,
      options: MarkdownOptions = {}
    ): string {
      return deps.createOutputFileName(
        workbookName,
        sheetIndex,
        sheetName,
        options.outputMode || "display",
        options.formattingMode || "plain"
      ).replace(/\.md$/i, "");
    }

    function parseInternalHyperlinkLocation(location: string, currentSheetName: string): { sheetName: string; refText: string } {
      const normalized = String(location || "").trim().replace(/^#/, "");
      if (!normalized) {
        return { sheetName: currentSheetName, refText: "" };
      }
      const match = normalized.match(/^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/);
      if (match) {
        return {
          sheetName: (match[1] || match[2] || currentSheetName).replace(/''/g, "'"),
          refText: (match[3] || "").trim()
        };
      }
      return {
        sheetName: currentSheetName,
        refText: normalized
      };
    }

    function renderHyperlinkMarkdown(
      cell: ParsedCell,
      text: string,
      workbook: ParsedWorkbook | null,
      sheet: ParsedSheet | null,
      options: MarkdownOptions
    ): string {
      const hyperlink = cell.hyperlink;
      const label = String(text || "").trim();
      if (!hyperlink || !label) return text;
      if (hyperlink.kind === "external") {
        const href = String(hyperlink.target || "").trim();
        return href ? `[${label}](${href})` : label;
      }
      const currentSheetName = sheet?.name || "";
      const { sheetName, refText } = parseInternalHyperlinkLocation(hyperlink.location || hyperlink.target, currentSheetName);
      const traceText = [sheetName, refText].filter(Boolean).join("!");
      const targetSheet = workbook?.sheets.find((entry) => entry.name === sheetName) || null;
      if (!targetSheet || !workbook) {
        return traceText ? `${label} (${traceText})` : label;
      }
      const href = `#${createSheetAnchorId(workbook.name, targetSheet.index, targetSheet.name, options)}`;
      return traceText && traceText !== targetSheet.name
        ? `[${label}](${href}) (${traceText})`
        : `[${label}](${href})`;
    }

    function formatCellForMarkdown(
      cell: ParsedCell | undefined,
      options: MarkdownOptions,
      workbook: ParsedWorkbook | null = null,
      sheet: ParsedSheet | null = null
    ): string {
      if (!cell) return "";
      const mode = options.outputMode || "display";
      const formattingMode = options.formattingMode || "plain";
      const displayCell = formattingMode === "github" && cell.hyperlink
        ? {
          ...cell,
          textStyle: {
            ...cell.textStyle,
            underline: false
          },
          richTextRuns: cell.richTextRuns?.map((run) => ({
            ...run,
            underline: false
          })) || null
        }
        : cell;
      const displayValue = richTextRenderer.compactText(String(cell.outputValue || ""));
      const rawValue = richTextRenderer.compactText(String(cell.rawValue || ""));
      const displayMarkdown = richTextRenderer.renderCellDisplayText(displayCell, formattingMode);
      if (mode === "raw") {
        return renderHyperlinkMarkdown(cell, rawValue || displayValue, workbook, sheet, options);
      }
      if (mode === "both") {
        if (rawValue && rawValue !== displayValue) {
          if (displayMarkdown) {
            return `${renderHyperlinkMarkdown(cell, displayMarkdown, workbook, sheet, options)} [raw=${rawValue}]`;
          }
          return `[raw=${rawValue}]`;
        }
        return renderHyperlinkMarkdown(cell, displayMarkdown || rawValue, workbook, sheet, options);
      }
      return renderHyperlinkMarkdown(cell, displayMarkdown, workbook, sheet, options);
    }

    function isCellInAnyTable(row: number, col: number, tables: TableCandidate[]): boolean {
      return tables.some((table) => row >= table.startRow && row <= table.endRow && col >= table.startCol && col <= table.endCol);
    }

    function splitNarrativeRowSegments(
      cells: ParsedCell[],
      options: MarkdownOptions,
      workbook: ParsedWorkbook | null = null,
      sheet: ParsedSheet | null = null
    ): Array<{ startCol: number; values: string[] }> {
      const segments: Array<{ startCol: number; values: string[] }> = [];
      let current: { startCol: number; values: string[]; lastCol: number } | null = null;
      for (const cell of cells) {
        const value = formatCellForMarkdown(cell, options, workbook, sheet).trim();
        if (!value) continue;
        if (!current || cell.col - current.lastCol > 4) {
          current = {
            startCol: cell.col,
            values: [value],
            lastCol: cell.col
          };
          segments.push(current);
        } else {
          current.values.push(value);
          current.lastCol = cell.col;
        }
      }
      return segments.map((segment) => ({
        startCol: segment.startCol,
        values: segment.values
      }));
    }

    function extractNarrativeBlocks(
      workbook: ParsedWorkbook,
      sheet: ParsedSheet,
      tables: TableCandidate[],
      options: MarkdownOptions = {}
    ): NarrativeBlock[] {
      const rowMap = new Map<number, ParsedCell[]>();
      for (const cell of sheet.cells) {
        if (!cell.outputValue) continue;
        if (isCellInAnyTable(cell.row, cell.col, tables)) continue;
        const entries = rowMap.get(cell.row) || [];
        entries.push(cell);
        rowMap.set(cell.row, entries);
      }
      const rowNumbers = Array.from(rowMap.keys()).sort((a, b) => a - b);
      const blocks: NarrativeBlock[] = [];
      let current: NarrativeBlock | null = null;
      let previousRow = -100;

      for (const rowNumber of rowNumbers) {
        const cells = (rowMap.get(rowNumber) || []).slice().sort((a, b) => a.col - b.col);
        const rowSegments = splitNarrativeRowSegments(cells, options, workbook, sheet);
        for (const segment of rowSegments) {
          const rowText = segment.values.join(" ").trim();
          if (!rowText) continue;
          const startCol = segment.startCol;
          if (!current || rowNumber - previousRow > 1 || Math.abs(startCol - current.startCol) > 3) {
            current = {
              startRow: rowNumber,
              startCol,
              endRow: rowNumber,
              lines: [rowText],
              items: [{
                row: rowNumber,
                startCol,
                text: rowText,
                cellValues: segment.values
              }]
            };
            blocks.push(current);
          } else {
            current.lines.push(rowText);
            current.endRow = rowNumber;
            current.items.push({
              row: rowNumber,
              startCol,
              text: rowText,
              cellValues: segment.values
            });
          }
          previousRow = rowNumber;
        }
      }

      return blocks;
    }

    function extractSectionBlocks(sheet: ParsedSheet, tables: TableCandidate[], narrativeBlocks: NarrativeBlock[]): SectionBlock[] {
      const charts = sheet.charts || [];
      const anchors: Array<{ startRow: number; startCol: number; endRow: number; endCol: number }> = [];

      for (const block of narrativeBlocks) {
        anchors.push({
          startRow: block.startRow,
          startCol: block.startCol,
          endRow: block.endRow,
          endCol: Math.max(block.startCol, ...block.items.map((item) => item.startCol))
        });
      }

      for (const table of tables) {
        anchors.push({
          startRow: table.startRow,
          startCol: table.startCol,
          endRow: table.endRow,
          endCol: table.endCol
        });
      }

      for (const image of sheet.images) {
        const anchor = deps.parseCellAddress(image.anchor);
        if (anchor.row > 0 && anchor.col > 0) {
          anchors.push({ startRow: anchor.row, startCol: anchor.col, endRow: anchor.row, endCol: anchor.col });
        }
      }

      for (const chart of charts) {
        const anchor = deps.parseCellAddress(chart.anchor);
        if (anchor.row > 0 && anchor.col > 0) {
          anchors.push({ startRow: anchor.row, startCol: anchor.col, endRow: anchor.row, endCol: anchor.col });
        }
      }

      if (anchors.length === 0) {
        return [];
      }
      anchors.sort((left, right) => {
        if (left.startRow !== right.startRow) return left.startRow - right.startRow;
        return left.startCol - right.startCol;
      });

      const sections: SectionBlock[] = [];
      let current: SectionBlock | null = null;
      let previousEndRow = -100;
      const verticalGapThreshold = 4;

      for (const anchor of anchors) {
        const gap = anchor.startRow - previousEndRow;
        if (!current || gap > verticalGapThreshold) {
          current = {
            startRow: anchor.startRow,
            startCol: anchor.startCol,
            endRow: anchor.endRow,
            endCol: anchor.endCol
          };
          sections.push(current);
        } else {
          current.startRow = Math.min(current.startRow, anchor.startRow);
          current.startCol = Math.min(current.startCol, anchor.startCol);
          current.endRow = Math.max(current.endRow, anchor.endRow);
          current.endCol = Math.max(current.endCol, anchor.endCol);
        }
        previousEndRow = Math.max(previousEndRow, anchor.endRow);
      }

      return sections;
    }

    function convertSheetToMarkdown(workbook: ParsedWorkbook, sheet: ParsedSheet, options: MarkdownOptions = {}): MarkdownFile {
      const charts = sheet.charts || [];
      const shapes = sheet.shapes || [];
      const shapeBlocks = deps.extractShapeBlocks(shapes, {
        defaultCellWidthEmu: deps.defaultCellWidthEmu,
        defaultCellHeightEmu: deps.defaultCellHeightEmu,
        shapeBlockGapXEmu: deps.shapeBlockGapXEmu,
        shapeBlockGapYEmu: deps.shapeBlockGapYEmu
      });
      const treatFirstRowAsHeader = options.treatFirstRowAsHeader !== false;
      const tableDetectionMode = options.tableDetectionMode || "balanced";
      const tables = deps.detectTableCandidates(sheet, buildCellMap, tableDetectionMode);
      const narrativeBlocks = extractNarrativeBlocks(workbook, sheet, tables, options);
      const sectionBlocks = extractSectionBlocks(sheet, tables, narrativeBlocks);
      const formulaDiagnostics = sheet.cells
        .filter((cell) => !!cell.formulaText && cell.resolutionStatus !== null)
        .map((cell) => ({
          address: cell.address,
          formulaText: cell.formulaText,
          status: cell.resolutionStatus,
          source: cell.resolutionSource,
          outputValue: cell.outputValue
        }));
      const sections: Array<{
        sortRow: number;
        sortCol: number;
        markdown: string;
        kind: "narrative" | "table";
        narrativeBlock?: NarrativeBlock;
      }> = [];

      for (const block of narrativeBlocks) {
        sections.push({
          sortRow: block.startRow,
          sortCol: block.startCol,
          markdown: `${deps.renderNarrativeBlock(block)}\n`,
          kind: "narrative",
          narrativeBlock: block
        });
      }

      const fileName = deps.createOutputFileName(
        workbook.name,
        sheet.index,
        sheet.name,
        options.outputMode || "display",
        options.formattingMode || "plain"
      );
      const sheetAnchorId = createSheetAnchorId(workbook.name, sheet.index, sheet.name, options);
      let tableCounter = 1;
      for (const table of tables) {
        const rows = deps.matrixFromCandidate(
          sheet,
          table,
          options,
          buildCellMap,
          (cell, tableOptions) => formatCellForMarkdown(cell, tableOptions, workbook, sheet)
        );
        if (rows.length === 0 || rows[0]?.length === 0) continue;
        const tableMarkdown = deps.renderMarkdownTable(rows, treatFirstRowAsHeader);
        sections.push({
          sortRow: table.startRow,
          sortCol: table.startCol,
          markdown: `### Table ${String(tableCounter).padStart(3, "0")} (${deps.formatRange(table.startRow, table.startCol, table.endRow, table.endCol)})\n\n${tableMarkdown}\n`,
          kind: "table"
        });
        tableCounter += 1;
      }

      sections.sort((left, right) => {
        if (left.sortRow !== right.sortRow) return left.sortRow - right.sortRow;
        return left.sortCol - right.sortCol;
      });
      const groupedSections = (sectionBlocks.length > 0 ? sectionBlocks : [{
        startRow: -1,
        startCol: -1,
        endRow: Number.MAX_SAFE_INTEGER,
        endCol: Number.MAX_SAFE_INTEGER
      }]).map((block) => ({
        block,
        entries: sections.filter((section) =>
          section.sortRow >= block.startRow
          && section.sortRow <= block.endRow
          && section.sortCol >= block.startCol
          && section.sortCol <= block.endCol
        )
      })).filter((group) => group.entries.length > 0);

      const body = groupedSections
        .map((group) => group.entries.map((section) => section.markdown.trimEnd()).join("\n\n").trim())
        .filter(Boolean)
        .join("\n\n---\n\n")
        .trim();
      const imageSection = sheet.images.length > 0
        ? [
          "",
          "## Images",
          "",
          ...sheet.images.map((image, index) => [
            `### Image ${String(index + 1).padStart(3, "0")} (${image.anchor})`,
            `- File: ${image.path}`,
            "",
            `![${image.filename}](${image.path})`
          ].join("\n"))
        ].join("\n")
        : "";
      const chartSection = charts.length > 0
        ? [
          "",
          "## Charts",
          "",
          ...charts.map((chart, index) => {
            const lines = [
              `### Chart ${String(index + 1).padStart(3, "0")} (${chart.anchor})`,
              `- Title: ${chart.title || "(none)"}`,
              `- Type: ${chart.chartType}`
            ];
            if (chart.series.length > 0) {
              lines.push("- Series:");
              for (const series of chart.series) {
                lines.push(`  - ${series.name}`);
                if (series.axis === "secondary") lines.push("    - Axis: secondary");
                if (series.categoriesRef) lines.push(`    - categories: ${series.categoriesRef}`);
                if (series.valuesRef) lines.push(`    - values: ${series.valuesRef}`);
              }
            }
            return lines.join("\n");
          })
        ].join("\n")
        : "";
      const includeShapeDetails = options.includeShapeDetails !== false;
      const shapeSection = includeShapeDetails && shapes.length > 0
        ? [
          "",
          "## Shape Blocks",
          "",
          ...shapeBlocks.map((block, blockIndex) => [
            `### Shape Block ${String(blockIndex + 1).padStart(3, "0")} (${deps.formatRange(block.startRow, block.startCol, block.endRow, block.endCol)})`,
            `- Shapes: ${block.shapeIndexes.map((shapeIndex) => `Shape ${String(shapeIndex + 1).padStart(3, "0")}`).join(", ")}`,
            `- anchorRange: ${deps.colToLetters(block.startCol)}${block.startRow}-${deps.colToLetters(block.endCol)}${block.endRow}`
          ].join("\n")),
          "",
          "## Shapes",
          "",
          ...shapes.map((shape, index) => {
            const lines = [
              `### Shape ${String(index + 1).padStart(3, "0")} (${shape.anchor})`,
              ...deps.renderHierarchicalRawEntries(shape.rawEntries)
            ];
            if (shape.svgPath) {
              lines.push(`- SVG: ${shape.svgPath}`);
              lines.push("");
              lines.push(`![${shape.svgFilename || `shape_${String(index + 1).padStart(3, "0")}.svg`}](${shape.svgPath})`);
            }
            return lines.join("\n");
          })
        ].join("\n")
        : "";
      const markdown = [
        `<a id="${sheetAnchorId}"></a>`,
        "",
        `# ${sheet.name}`,
        "",
        "## Source Information",
        `- Workbook: ${workbook.name}`,
        `- Sheet: ${sheet.name}`,
        "",
        "## Body",
        "",
        body || "_No extractable body content was found._",
        chartSection,
        shapeSection,
        imageSection
      ].join("\n");

      return {
        fileName,
        sheetName: sheet.name,
        markdown,
        summary: {
          outputMode: options.outputMode || "display",
          formattingMode: options.formattingMode || "plain",
          tableDetectionMode,
          sections: sectionBlocks.length,
          tables: tables.length,
          narrativeBlocks: narrativeBlocks.length,
          merges: sheet.merges.length,
          images: sheet.images.length,
          charts: charts.length,
          cells: sheet.cells.length,
          tableScores: tables.map((table) => ({
            range: deps.formatRange(table.startRow, table.startCol, table.endRow, table.endCol),
            score: table.score,
            reasons: [...table.reasonSummary]
          })),
          formulaDiagnostics
        }
      };
    }

    function convertWorkbookToMarkdownFiles(workbook: ParsedWorkbook, options: MarkdownOptions = {}): MarkdownFile[] {
      return workbook.sheets.map((sheet) => convertSheetToMarkdown(workbook, sheet, options));
    }

    return {
      buildCellMap,
      formatCellForMarkdown,
      isCellInAnyTable,
      splitNarrativeRowSegments,
      extractNarrativeBlocks,
      extractSectionBlocks,
      convertSheetToMarkdown,
      convertWorkbookToMarkdownFiles
    };
  }

  const sheetMarkdownApi = {
    createSheetMarkdownApi
  };

  moduleRegistry.registerModule("sheetMarkdown", sheetMarkdownApi);
})();
