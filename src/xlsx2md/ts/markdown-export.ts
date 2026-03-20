(() => {
  const moduleRegistry = getXlsx2mdModuleRegistry();
  type FormulaResolutionStatus = "resolved" | "fallback_formula" | "unsupported_external" | null;

  type FormulaDiagnostic = {
    address: string;
    formulaText: string;
    status: FormulaResolutionStatus;
    source: string | null;
    outputValue: string;
  };

  type TableScoreDetail = {
    range: string;
    score: number;
    reasons: string[];
  };

  type MarkdownFile = {
    fileName: string;
    sheetName: string;
    markdown: string;
    summary: {
      outputMode: "display" | "raw" | "both";
      sections: number;
      tables: number;
      narrativeBlocks: number;
      merges: number;
      images: number;
      charts: number;
      cells: number;
      tableScores: TableScoreDetail[];
      formulaDiagnostics: FormulaDiagnostic[];
    };
  };

  type ExportEntry = { name: string; data: Uint8Array };

  type WorkbookLike = {
    name: string;
    sheets: Array<{
      images: Array<{
        path: string;
        data: Uint8Array;
      }>;
      shapes?: Array<{
        svgPath: string | null;
        svgData: Uint8Array | null;
      }>;
    }>;
  };

  const textEncoder = new TextEncoder();
  const zipIoHelper = requireXlsx2mdZipIo();
  const markdownNormalizeHelper = requireXlsx2mdMarkdownNormalize();

  function normalizeMarkdownLineBreaks(text: string): string {
    return markdownNormalizeHelper.normalizeMarkdownText(text);
  }

  function escapeMarkdownCell(text: string): string {
    return markdownNormalizeHelper.normalizeMarkdownTableCell(text);
  }

  function renderMarkdownTable(rows: string[][], treatFirstRowAsHeader: boolean): string {
    if (rows.length === 0) {
      return "";
    }
    const workingRows = rows.map((row) => row.map((cell) => escapeMarkdownCell(cell)));
    if (workingRows.length === 1 && treatFirstRowAsHeader) {
      workingRows.push(new Array(workingRows[0].length).fill(""));
    }
    const header = treatFirstRowAsHeader ? workingRows[0] : new Array(workingRows[0].length).fill("");
    const body = treatFirstRowAsHeader ? workingRows.slice(1) : workingRows;
    const lines = [
      `| ${header.join(" | ")} |`,
      `| ${header.map(() => "---").join(" | ")} |`
    ];
    for (const row of body) {
      lines.push(`| ${row.join(" | ")} |`);
    }
    return lines.join("\n");
  }

  function sanitizeFileNameSegment(value: string, fallback: string): string {
    const normalized = String(value || "").normalize("NFKC");
    const sanitized = normalized
      .replace(/[\\/:*?"<>|]/g, "_")
      .replace(/\s+/g, "_")
      .replace(/[^\p{L}\p{N}._-]+/gu, "_")
      .replace(/_+/g, "_")
      .replace(/^[_ .-]+|[_ .-]+$/g, "");
    return sanitized || fallback;
  }

  function createOutputFileName(
    workbookName: string,
    sheetIndex: number,
    sheetName: string,
    outputMode: "display" | "raw" | "both" = "display"
  ): string {
    const bookBase = sanitizeFileNameSegment(workbookName.replace(/\.xlsx$/i, ""), "workbook");
    const safeSheetName = sanitizeFileNameSegment(sheetName, `Sheet${sheetIndex}`);
    const suffix = outputMode === "display" ? "" : `_${outputMode}`;
    return `${bookBase}_${String(sheetIndex).padStart(3, "0")}_${safeSheetName}${suffix}.md`;
  }

  function createSummaryText(markdownFile: MarkdownFile): string {
    const resolvedCount = markdownFile.summary.formulaDiagnostics.filter((item) => item.status === "resolved").length;
    const fallbackCount = markdownFile.summary.formulaDiagnostics.filter((item) => item.status === "fallback_formula").length;
    const unsupportedCount = markdownFile.summary.formulaDiagnostics.filter((item) => item.status === "unsupported_external").length;
    return [
      `Output file: ${markdownFile.fileName}`,
      `Output mode: ${markdownFile.summary.outputMode}`,
      `Sections: ${markdownFile.summary.sections}`,
      `Tables: ${markdownFile.summary.tables}`,
      `Narrative blocks: ${markdownFile.summary.narrativeBlocks}`,
      `Merged ranges: ${markdownFile.summary.merges}`,
      `Images: ${markdownFile.summary.images}`,
      `Charts: ${markdownFile.summary.charts}`,
      `Analyzed cells: ${markdownFile.summary.cells}`,
      `Formula resolved: ${resolvedCount}`,
      `Formula fallback_formula: ${fallbackCount}`,
      `Formula unsupported_external: ${unsupportedCount}`,
      ...markdownFile.summary.tableScores.map((detail) => `Table candidate ${detail.range}: score ${detail.score} / ${detail.reasons.join(", ")}`)
    ].join("\n");
  }

  function createCombinedMarkdownExportFile(workbook: WorkbookLike, markdownFiles: MarkdownFile[]): { fileName: string; content: string } {
    const outputMode = markdownFiles[0]?.summary.outputMode || "display";
    const suffix = outputMode === "display" ? "" : `_${outputMode}`;
    const fileName = `${String(workbook.name || "workbook").replace(/\.xlsx$/i, "")}${suffix}.md`;
    const content = markdownFiles
      .map((markdownFile) => `<!-- ${markdownFile.fileName.replace(/\.md$/i, "")} -->\n${markdownFile.markdown}`)
      .join("\n\n");
    return { fileName, content };
  }

  function createExportEntries(workbook: WorkbookLike, markdownFiles: MarkdownFile[]): ExportEntry[] {
    const entries: ExportEntry[] = [];
    if (markdownFiles.length > 0) {
      const combined = createCombinedMarkdownExportFile(workbook, markdownFiles);
      entries.push({
        name: `output/${combined.fileName}`,
        data: textEncoder.encode(`${combined.content}\n`)
      });
    }
    for (const sheet of workbook.sheets) {
      for (const image of sheet.images) {
        entries.push({
          name: `output/${image.path}`,
          data: image.data
        });
      }
      for (const shape of sheet.shapes || []) {
        if (!shape.svgPath || !shape.svgData) continue;
        entries.push({
          name: `output/${shape.svgPath}`,
          data: shape.svgData
        });
      }
    }
    return entries;
  }

  function createWorkbookExportArchive(workbook: WorkbookLike, markdownFiles: MarkdownFile[]): Uint8Array {
    return zipIoHelper.createStoredZip(createExportEntries(workbook, markdownFiles));
  }

  const markdownExportApi = {
    escapeMarkdownCell,
    renderMarkdownTable,
    sanitizeFileNameSegment,
    createOutputFileName,
    createSummaryText,
    createCombinedMarkdownExportFile,
    createExportEntries,
    createWorkbookExportArchive,
    normalizeMarkdownLineBreaks,
    textEncoder
  };

  moduleRegistry.registerModule("markdownExport", markdownExportApi);
})();
