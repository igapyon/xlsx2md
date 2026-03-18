(() => {
  type MarkdownOptions = {
    treatFirstRowAsHeader: boolean;
    trimText: boolean;
    removeEmptyRows: boolean;
    removeEmptyColumns: boolean;
    outputMode: "display" | "raw" | "both";
  };

  type WorkbookFile = {
    fileName: string;
    sheetName: string;
    markdown: string;
    summary: {
      outputMode: "display" | "raw" | "both";
      tables: number;
      narrativeBlocks: number;
      merges: number;
      images: number;
      cells: number;
      tableScores: Array<{
        range: string;
        score: number;
        reasons: string[];
      }>;
      formulaDiagnostics: Array<{
        address: string;
        formulaText: string;
        status: "resolved" | "fallback_formula" | "unsupported_external" | null;
        source: "cached_value" | "ast_evaluator" | "legacy_resolver" | "formula_text" | "external_unsupported" | null;
        outputValue: string;
      }>;
    };
  };

  type ParsedWorkbook = {
    name: string;
    sheets: Array<{ name: string; index: number }>;
  };

  const xlsx2md = (globalThis as typeof globalThis & {
    __xlsx2md?: {
      parseWorkbook: (arrayBuffer: ArrayBuffer, workbookName?: string) => Promise<ParsedWorkbook & { sheets: Array<Record<string, unknown>> }>;
      convertWorkbookToMarkdownFiles: (workbook: ParsedWorkbook & { sheets: Array<Record<string, unknown>> }, options?: MarkdownOptions) => WorkbookFile[];
      createSummaryText: (file: WorkbookFile) => string;
      createCombinedMarkdownExportFile: (workbook: ParsedWorkbook & { sheets: Array<Record<string, unknown>> }, files: WorkbookFile[]) => { fileName: string; content: string };
      createWorkbookExportArchive: (workbook: ParsedWorkbook & { sheets: Array<Record<string, unknown>> }, files: WorkbookFile[]) => Uint8Array;
    };
  }).__xlsx2md;

  if (!xlsx2md) {
    throw new Error("xlsx2md core module is not loaded");
  }

  let currentWorkbook: (ParsedWorkbook & { sheets: Array<Record<string, unknown>> }) | null = null;
  let currentFiles: WorkbookFile[] = [];

  function getElement<T extends HTMLElement>(id: string): T {
    const element = document.getElementById(id);
    if (!element) {
      throw new Error(`Element not found: ${id}`);
    }
    return element as T;
  }

  function getSwitchValue(id: string): boolean {
    const element = getElement<HTMLInputElement>(id);
    return !!element.checked;
  }

  function getOptions(): MarkdownOptions {
    const outputModeSelect = getElement<HTMLElement>("outputModeSelect") as HTMLElement & { getValue?: () => string };
    const outputMode = typeof outputModeSelect.getValue === "function"
      ? outputModeSelect.getValue()
      : (document.getElementById("outputModeSelect") as HTMLSelectElement | null)?.value || "display";
    return {
      treatFirstRowAsHeader: getSwitchValue("headerRowEnabled"),
      trimText: getSwitchValue("trimTextEnabled"),
      removeEmptyRows: getSwitchValue("removeEmptyRowsEnabled"),
      removeEmptyColumns: getSwitchValue("removeEmptyColumnsEnabled"),
      outputMode: outputMode === "raw" || outputMode === "both" ? outputMode : "display"
    };
  }

  function getSelectedOutputMode(): "display" | "raw" | "both" {
    return getOptions().outputMode;
  }

  function showToast(message: string): void {
    const toast = document.getElementById("toast") as (HTMLElement & { show?: (text: string, duration?: number) => void }) | null;
    if (toast && typeof toast.show === "function") {
      toast.show(message, 2200);
    }
  }

  function setSummaryHtml(html: string): void {
    getElement<HTMLElement>("analysisSummary").innerHTML = html;
  }

  function setSummaryText(message: string): void {
    setSummaryHtml(`<div class="md-summary-empty">${escapeHtml(message)}</div>`);
  }

  function setScoreSummaryHtml(html: string): void {
    getElement<HTMLElement>("scoreSummary").innerHTML = html;
  }

  function setFormulaSummaryHtml(html: string): void {
    getElement<HTMLElement>("formulaSummary").innerHTML = html;
  }

  function escapeHtml(value: string): string {
    return String(value || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function renderScoreSummary(files: WorkbookFile[]): string {
    const sheetsWithScores = files.filter((file) => file.summary.tableScores.length > 0);
    if (sheetsWithScores.length === 0) {
      return '<div class="md-summary-empty">表候補はありません。</div>';
    }
    const totalScores = sheetsWithScores.reduce((sum, file) => sum + file.summary.tableScores.length, 0);
    const totalStrong = sheetsWithScores.reduce((sum, file) => (
      sum + file.summary.tableScores.filter((detail) => getTableScoreLabel(detail.score) === "strong").length
    ), 0);
    const totalCandidate = sheetsWithScores.reduce((sum, file) => (
      sum + file.summary.tableScores.filter((detail) => getTableScoreLabel(detail.score) === "candidate").length
    ), 0);
    return `<div class="md-summary-overview">全体 ${totalScores}件 / strong ${totalStrong} / candidate ${totalCandidate}</div>${sheetsWithScores.map((file) => {
      const items = [...file.summary.tableScores].sort((left, right) => {
        if (right.score !== left.score) {
          return right.score - left.score;
        }
        return left.range.localeCompare(right.range);
      }).map((detail) => (
        `<div class="md-summary-item"><div class="md-summary-item-head"><span class="md-summary-item-title">${escapeHtml(detail.range)}</span><span class="md-summary-item-status md-summary-item-status--${escapeHtml(getTableScoreLabel(detail.score))}">${escapeHtml(getTableScoreText(detail.score))}</span></div><div class="md-summary-item-meta">${detail.score}点</div><div class="md-summary-item-body">${escapeHtml(detail.reasons.join(" / "))}</div></div>`
      )).join("");
      return `<section class="md-summary-group"><div class="md-summary-group-head"><h3 class="md-summary-group-title">${escapeHtml(file.sheetName)}</h3><span class="md-summary-group-count">${file.summary.tableScores.length}件</span></div><div class="md-summary-group-meta">${escapeHtml(renderTableScoreCounts(file))}</div>${items}</section>`;
    }).join("")}`;
  }

  function getTableScoreLabel(score: number): string {
    if (score >= 7) return "strong";
    if (score >= 4) return "candidate";
    return "unknown";
  }

  function getTableScoreText(score: number): string {
    if (score >= 7) return "strong";
    if (score >= 4) return "candidate";
    return "unknown";
  }

  function renderTableScoreCounts(file: WorkbookFile): string {
    const counts = {
      strong: 0,
      candidate: 0,
      unknown: 0
    };
    file.summary.tableScores.forEach((detail) => {
      counts[getTableScoreLabel(detail.score) as keyof typeof counts] += 1;
    });
    return [
      counts.strong > 0 ? `strong ${counts.strong}` : "",
      counts.candidate > 0 ? `candidate ${counts.candidate}` : "",
      counts.unknown > 0 ? `unknown ${counts.unknown}` : ""
    ].filter(Boolean).join(" / ");
  }

  function getFormulaStatusLabel(status: "resolved" | "fallback_formula" | "unsupported_external" | null): string {
    if (status === "resolved") return "resolved";
    if (status === "fallback_formula") return "fallback";
    if (status === "unsupported_external") return "unsupported";
    return "unknown";
  }

  function renderFormulaStatusCounts(file: WorkbookFile): string {
    const counts = {
      resolved: 0,
      fallback: 0,
      unsupported: 0,
      unknown: 0
    };
    file.summary.formulaDiagnostics.forEach((diagnostic) => {
      counts[getFormulaStatusLabel(diagnostic.status) as keyof typeof counts] += 1;
    });
    return [
      counts.resolved > 0 ? `resolved ${counts.resolved}` : "",
      counts.fallback > 0 ? `fallback ${counts.fallback}` : "",
      counts.unsupported > 0 ? `unsupported ${counts.unsupported}` : "",
      counts.unknown > 0 ? `unknown ${counts.unknown}` : ""
    ].filter(Boolean).join(" / ");
  }

  function getFormulaStatusPriority(status: "resolved" | "fallback_formula" | "unsupported_external" | null): number {
    const label = getFormulaStatusLabel(status);
    if (label === "unsupported") return 0;
    if (label === "fallback") return 1;
    if (label === "unknown") return 2;
    return 3;
  }

  function getFormulaSourceLabel(source: "cached_value" | "ast_evaluator" | "legacy_resolver" | "formula_text" | "external_unsupported" | null): string {
    if (source === "cached_value") return "cached";
    if (source === "ast_evaluator") return "ast";
    if (source === "legacy_resolver") return "legacy";
    if (source === "formula_text") return "formula";
    if (source === "external_unsupported") return "external";
    return "unknown";
  }

  function renderFormulaSourceCounts(file: WorkbookFile): string {
    const counts = {
      cached: 0,
      ast: 0,
      legacy: 0,
      formula: 0,
      external: 0,
      unknown: 0
    };
    file.summary.formulaDiagnostics.forEach((diagnostic) => {
      counts[getFormulaSourceLabel(diagnostic.source) as keyof typeof counts] += 1;
    });
    return [
      counts.cached > 0 ? `cached ${counts.cached}` : "",
      counts.ast > 0 ? `ast ${counts.ast}` : "",
      counts.legacy > 0 ? `legacy ${counts.legacy}` : "",
      counts.formula > 0 ? `formula ${counts.formula}` : "",
      counts.external > 0 ? `external ${counts.external}` : "",
      counts.unknown > 0 ? `unknown ${counts.unknown}` : ""
    ].filter(Boolean).join(" / ");
  }

  function renderFormulaSummary(files: WorkbookFile[]): string {
    const sheetsWithDiagnostics = files.filter((file) => file.summary.formulaDiagnostics.length > 0);
    if (sheetsWithDiagnostics.length === 0) {
      return '<div class="md-summary-empty">数式セルはありません。</div>';
    }
    const totalDiagnostics = sheetsWithDiagnostics.reduce((sum, file) => sum + file.summary.formulaDiagnostics.length, 0);
    const totalResolved = sheetsWithDiagnostics.reduce((sum, file) => (
      sum + file.summary.formulaDiagnostics.filter((diagnostic) => getFormulaStatusLabel(diagnostic.status) === "resolved").length
    ), 0);
    const totalFallback = sheetsWithDiagnostics.reduce((sum, file) => (
      sum + file.summary.formulaDiagnostics.filter((diagnostic) => getFormulaStatusLabel(diagnostic.status) === "fallback").length
    ), 0);
    const totalUnsupported = sheetsWithDiagnostics.reduce((sum, file) => (
      sum + file.summary.formulaDiagnostics.filter((diagnostic) => getFormulaStatusLabel(diagnostic.status) === "unsupported").length
    ), 0);
    const totalCached = sheetsWithDiagnostics.reduce((sum, file) => (
      sum + file.summary.formulaDiagnostics.filter((diagnostic) => diagnostic.source === "cached_value").length
    ), 0);
    const totalAst = sheetsWithDiagnostics.reduce((sum, file) => (
      sum + file.summary.formulaDiagnostics.filter((diagnostic) => diagnostic.source === "ast_evaluator").length
    ), 0);
    const totalLegacy = sheetsWithDiagnostics.reduce((sum, file) => (
      sum + file.summary.formulaDiagnostics.filter((diagnostic) => diagnostic.source === "legacy_resolver").length
    ), 0);
    const totalFormula = sheetsWithDiagnostics.reduce((sum, file) => (
      sum + file.summary.formulaDiagnostics.filter((diagnostic) => diagnostic.source === "formula_text").length
    ), 0);
    return `<div class="md-summary-overview">全体 ${totalDiagnostics}件 / cached ${totalCached} / ast ${totalAst} / legacy ${totalLegacy} / formula ${totalFormula} / resolved ${totalResolved} / fallback ${totalFallback} / unsupported ${totalUnsupported}</div>${sheetsWithDiagnostics.map((file) => {
      const items = [...file.summary.formulaDiagnostics].sort((left, right) => {
        const priorityDiff = getFormulaStatusPriority(left.status) - getFormulaStatusPriority(right.status);
        if (priorityDiff !== 0) {
          return priorityDiff;
        }
        return left.address.localeCompare(right.address);
      }).map((diagnostic) => (
        `<div class="md-summary-item"><div class="md-summary-item-head"><span class="md-summary-item-title">${escapeHtml(diagnostic.address)}</span><span class="md-summary-item-badges"><span class="md-summary-item-status md-summary-item-status--source-${escapeHtml(getFormulaSourceLabel(diagnostic.source))}">${escapeHtml(getFormulaSourceLabel(diagnostic.source))}</span><span class="md-summary-item-status md-summary-item-status--${escapeHtml(getFormulaStatusLabel(diagnostic.status))}">${escapeHtml(getFormulaStatusLabel(diagnostic.status))}</span></span></div><div class="md-summary-item-body">${escapeHtml(`${diagnostic.formulaText} => ${diagnostic.outputValue}`)}</div></div>`
      )).join("");
      return `<section class="md-summary-group"><div class="md-summary-group-head"><h3 class="md-summary-group-title">${escapeHtml(file.sheetName)}</h3><span class="md-summary-group-count">${file.summary.formulaDiagnostics.length}件</span></div><div class="md-summary-group-meta">${escapeHtml(renderFormulaSourceCounts(file))}</div><div class="md-summary-group-meta">${escapeHtml(renderFormulaStatusCounts(file))}</div>${items}</section>`;
    }).join("")}`;
  }

  function renderAnalysisSummary(files: WorkbookFile[], workbookName: string): string {
    if (files.length === 0) {
      return '<div class="md-summary-empty">まだ変換していません。</div>';
    }
    const totalTables = files.reduce((sum, file) => sum + file.summary.tables, 0);
    const totalNarratives = files.reduce((sum, file) => sum + file.summary.narrativeBlocks, 0);
    const totalMerges = files.reduce((sum, file) => sum + file.summary.merges, 0);
    const totalImages = files.reduce((sum, file) => sum + file.summary.images, 0);
    const totalCells = files.reduce((sum, file) => sum + file.summary.cells, 0);
    const totalFormulas = files.reduce((sum, file) => sum + file.summary.formulaDiagnostics.length, 0);
    const outputMode = files[0]?.summary.outputMode || "display";
    const overview = `<div class="md-summary-overview">Workbook ${escapeHtml(workbookName)} / ${files.length} sheet / mode ${escapeHtml(outputMode)}</div>`;
    const items = files.map((file) => (
      `<section class="md-summary-group"><div class="md-summary-group-head"><h3 class="md-summary-group-title">${escapeHtml(file.sheetName)}</h3><span class="md-summary-group-count">${file.summary.cells} cells</span></div><div class="md-summary-group-meta">表 ${file.summary.tables} / 地の文 ${file.summary.narrativeBlocks} / 結合 ${file.summary.merges} / 画像 ${file.summary.images} / 数式 ${file.summary.formulaDiagnostics.length}</div></section>`
    )).join("");
    const totals = `<section class="md-summary-group"><div class="md-summary-group-head"><h3 class="md-summary-group-title">全体</h3><span class="md-summary-group-count">${files.length} sheets</span></div><div class="md-summary-group-meta">表 ${totalTables} / 地の文 ${totalNarratives} / 結合 ${totalMerges} / 画像 ${totalImages} / 数式 ${totalFormulas} / 解析セル ${totalCells}</div></section>`;
    return `${overview}${totals}${items}`;
  }

  function updateOutputModeNotice(mode: "display" | "raw" | "both"): void {
    const notice = getElement<HTMLElement>("outputModeNotice");
    if (mode === "raw") {
      notice.textContent = "`raw` は Excel の表示値ではなく、内部値を優先して Markdown に出力します。";
      return;
    }
    if (mode === "both") {
      notice.textContent = "`both` は表示値に加えて `[raw=...]` 形式の補助情報を出力します。";
      return;
    }
    notice.textContent = "`display` は Excel の表示値寄りで出力します。";
  }

  function updatePreviewModeBanner(mode: "display" | "raw" | "both"): void {
    const banner = getElement<HTMLElement>("previewModeBanner");
    if (mode === "raw") {
      banner.hidden = false;
      banner.textContent = "`raw` モードです。Markdown には Excel の表示値ではなく内部値が出ます。";
      return;
    }
    if (mode === "both") {
      banner.hidden = false;
      banner.textContent = "`both` モードです。Markdown には表示値に加えて `[raw=...]` が出ます。";
      return;
    }
    banner.hidden = true;
    banner.textContent = "";
  }

  function setPreviewMarkdown(markdown: string): void {
    const preview = getElement<HTMLElement>("markdownPreview") as HTMLElement & { setText?: (text: string) => void };
    if (typeof preview.setText === "function") {
      preview.setText(markdown);
      return;
    }
    getElement<HTMLElement>("markdownOutput").textContent = markdown;
  }

  function createMarkdownChunkLabel(fileName: string): string {
    return String(fileName || "").replace(/\.md$/i, "");
  }

  function clearError(): void {
    const errorAlert = getElement<HTMLElement>("errorAlert") as HTMLElement & { clear?: () => void };
    if (typeof errorAlert.clear === "function") {
      errorAlert.clear();
    } else {
      errorAlert.removeAttribute("active");
      errorAlert.textContent = "";
    }
  }

  function showError(message: string): void {
    const errorAlert = getElement<HTMLElement>("errorAlert") as HTMLElement & { show?: (text: string) => void };
    if (typeof errorAlert.show === "function") {
      errorAlert.show(message);
    } else {
      errorAlert.textContent = message;
      errorAlert.setAttribute("active", "");
    }
  }

  function setLoading(active: boolean, message?: string): void {
    const overlay = getElement<HTMLElement>("loadingOverlay") as HTMLElement & { show?: (text?: string) => void; hide?: () => void };
    if (active) {
      if (message) {
        overlay.setAttribute("text", message);
      }
      if (typeof overlay.show === "function") {
        overlay.show(message || "処理中です");
      } else {
        overlay.setAttribute("active", "");
      }
      return;
    }
    if (typeof overlay.hide === "function") {
      overlay.hide();
    } else {
      overlay.removeAttribute("active");
    }
  }

  function renderCurrentSelection(): void {
    if (!currentFiles.length) {
      setSummaryText("まだ変換していません。");
      setScoreSummaryHtml('<div class="md-summary-empty">まだ変換していません。</div>');
      setFormulaSummaryHtml('<div class="md-summary-empty">まだ変換していません。</div>');
      setPreviewMarkdown("");
      updatePreviewModeBanner(getSelectedOutputMode());
      return;
    }
    const combinedMarkdown = currentFiles
      .map((file) => `<!-- ${createMarkdownChunkLabel(file.fileName)} -->\n${file.markdown}`)
      .join("\n\n");
    const outputMode = currentFiles[0]?.summary.outputMode || "display";
    updatePreviewModeBanner(outputMode);
    setSummaryHtml(renderAnalysisSummary(currentFiles, currentWorkbook?.name || "workbook.xlsx"));
    setScoreSummaryHtml(renderScoreSummary(currentFiles));
    setFormulaSummaryHtml(renderFormulaSummary(currentFiles));
    setPreviewMarkdown(combinedMarkdown);
    getElement<HTMLButtonElement>("downloadBtn").disabled = false;
    getElement<HTMLButtonElement>("exportZipBtn").disabled = false;
  }

  function getSelectedFileForDownload(): { fileName: string; content: string } | null {
    if (!currentFiles.length) return null;
    if (!currentWorkbook) return null;
    return xlsx2md.createCombinedMarkdownExportFile(currentWorkbook, currentFiles);
  }

  function downloadCurrentMarkdown(): void {
    const payload = getSelectedFileForDownload();
    if (!payload) {
      showError("保存対象の Markdown がありません");
      return;
    }
    const blob = new Blob([`${payload.content}\n`], { type: "text/markdown;charset=utf-8" });
    const objectUrl = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = objectUrl;
    link.download = payload.fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    window.setTimeout(() => URL.revokeObjectURL(objectUrl), 0);
    showToast("Markdown を保存しました");
  }

  function downloadExportZip(): void {
    if (!currentWorkbook || currentFiles.length === 0) {
      showError("先に Markdown を生成してください");
      return;
    }
    const zipBytes = xlsx2md.createWorkbookExportArchive(currentWorkbook, currentFiles);
    const outputMode = currentFiles[0]?.summary.outputMode || "display";
    const suffix = outputMode === "display" ? "" : `_${outputMode}`;
    const blob = new Blob([zipBytes], { type: "application/zip" });
    const objectUrl = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = objectUrl;
    link.download = `${currentWorkbook.name.replace(/\.xlsx$/i, "")}_xlsx2md_export${suffix}.zip`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    window.setTimeout(() => URL.revokeObjectURL(objectUrl), 0);
    showToast("ZIP を保存しました");
  }

  function convertCurrentWorkbook(showSuccessToast = true): void {
    clearError();
    if (!currentWorkbook) {
      showError("先に xlsx ファイルを読み込んでください");
      return;
    }
    try {
      currentFiles = xlsx2md.convertWorkbookToMarkdownFiles(currentWorkbook, getOptions());
      renderCurrentSelection();
      if (showSuccessToast) {
        showToast("Markdown を生成しました");
      }
    } catch (error) {
      showError(error instanceof Error ? error.message : "Markdown 生成に失敗しました");
    }
  }

  async function loadWorkbookFromFile(file: File): Promise<void> {
    clearError();
    setLoading(true, "xlsx を読み込んでいます");
    try {
      const arrayBuffer = await file.arrayBuffer();
      currentWorkbook = await xlsx2md.parseWorkbook(arrayBuffer, file.name);
      currentFiles = [];
      convertCurrentWorkbook(false);
      showToast("xlsx を読み込み、Markdown を生成しました");
    } catch (error) {
      currentWorkbook = null;
      currentFiles = [];
      setSummaryText("Workbook の読込に失敗しました。");
      setScoreSummaryHtml('<div class="md-summary-empty">まだ変換していません。</div>');
      setFormulaSummaryHtml('<div class="md-summary-empty">まだ変換していません。</div>');
      setPreviewMarkdown("");
      getElement<HTMLButtonElement>("downloadBtn").disabled = true;
      getElement<HTMLButtonElement>("exportZipBtn").disabled = true;
      showError(error instanceof Error ? error.message : "xlsx の読込に失敗しました");
    } finally {
      setLoading(false);
    }
  }

  function bindFileInput(): void {
    const fileInput = getElement<HTMLInputElement>("xlsxFileInput");
    fileInput.addEventListener("change", async () => {
      const file = fileInput.files?.[0];
      if (!file) return;
      await loadWorkbookFromFile(file);
    });
  }

  function bindActions(): void {
    getElement<HTMLButtonElement>("convertBtn").addEventListener("click", () => {
      convertCurrentWorkbook(true);
    });
    getElement<HTMLButtonElement>("downloadBtn").addEventListener("click", () => {
      downloadCurrentMarkdown();
    });
    getElement<HTMLButtonElement>("exportZipBtn").addEventListener("click", () => {
      downloadExportZip();
    });
    getElement<HTMLElement>("outputModeSelect").addEventListener("change", () => {
      const mode = getSelectedOutputMode();
      updateOutputModeNotice(mode);
      if (!currentFiles.length) {
        updatePreviewModeBanner(mode);
      }
    });
  }

  function initialize(): void {
    clearError();
    setSummaryText("まだ変換していません。");
    setScoreSummaryHtml('<div class="md-summary-empty">まだ変換していません。</div>');
    setFormulaSummaryHtml('<div class="md-summary-empty">まだ変換していません。</div>');
    setPreviewMarkdown("");
    updateOutputModeNotice(getSelectedOutputMode());
    updatePreviewModeBanner(getSelectedOutputMode());
    getElement<HTMLButtonElement>("downloadBtn").disabled = true;
    getElement<HTMLButtonElement>("exportZipBtn").disabled = true;
    bindFileInput();
    bindActions();
  }

  document.addEventListener("DOMContentLoaded", initialize);
})();
