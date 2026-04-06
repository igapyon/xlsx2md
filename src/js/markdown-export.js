/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const moduleRegistry = getXlsx2mdModuleRegistry();
    const textEncoder = new TextEncoder();
    const zipIoHelper = requireXlsx2mdZipIo();
    const textEncodingHelper = requireXlsx2mdTextEncoding();
    const markdownNormalizeHelper = requireXlsx2mdMarkdownNormalize();
    const markdownTableEscapeHelper = requireXlsx2mdMarkdownTableEscape();
    function normalizeMarkdownLineBreaks(text) {
        return markdownNormalizeHelper.normalizeMarkdownText(text);
    }
    function escapeMarkdownCell(text) {
        return markdownTableEscapeHelper.escapeMarkdownTableCell(text);
    }
    function renderMarkdownTable(rows, treatFirstRowAsHeader) {
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
    function sanitizeFileNameSegment(value, fallback) {
        const normalized = String(value || "").normalize("NFKC");
        const sanitized = normalized
            .replace(/[\\/:*?"<>|]/g, "_")
            .replace(/\s+/g, "_")
            .replace(/[^\p{L}\p{N}._-]+/gu, "_")
            .replace(/_+/g, "_")
            .replace(/^[_ .-]+|[_ .-]+$/g, "");
        return sanitized || fallback;
    }
    function createOutputFileName(workbookName, sheetIndex, sheetName, outputMode = "display", formattingMode = "plain") {
        const bookBase = sanitizeFileNameSegment(workbookName.replace(/\.xlsx$/i, ""), "workbook");
        const safeSheetName = sanitizeFileNameSegment(sheetName, `Sheet${sheetIndex}`);
        const suffix = `${outputMode === "display" ? "" : `_${outputMode}`}${formattingMode === "plain" ? "" : `_${formattingMode}`}`;
        return `${bookBase}_${String(sheetIndex).padStart(3, "0")}_${safeSheetName}${suffix}.md`;
    }
    function createSummaryText(markdownFile) {
        const resolvedCount = markdownFile.summary.formulaDiagnostics.filter((item) => item.status === "resolved").length;
        const fallbackCount = markdownFile.summary.formulaDiagnostics.filter((item) => item.status === "fallback_formula").length;
        const unsupportedCount = markdownFile.summary.formulaDiagnostics.filter((item) => item.status === "unsupported_external").length;
        return [
            `Output file: ${markdownFile.fileName}`,
            `Output mode: ${markdownFile.summary.outputMode}`,
            `Formatting mode: ${markdownFile.summary.formattingMode}`,
            `Table detection mode: ${markdownFile.summary.tableDetectionMode}`,
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
    function createCombinedMarkdownExportFile(workbook, markdownFiles) {
        var _a, _b;
        const outputMode = ((_a = markdownFiles[0]) === null || _a === void 0 ? void 0 : _a.summary.outputMode) || "display";
        const formattingMode = ((_b = markdownFiles[0]) === null || _b === void 0 ? void 0 : _b.summary.formattingMode) || "plain";
        const suffix = `${outputMode === "display" ? "" : `_${outputMode}`}${formattingMode === "plain" ? "" : `_${formattingMode}`}`;
        const fileName = `${String(workbook.name || "workbook").replace(/\.xlsx$/i, "")}${suffix}.md`;
        const bookHeading = `# Book: ${String(workbook.name || "workbook.xlsx")}`;
        const content = [
            bookHeading,
            ...markdownFiles.map((markdownFile) => {
                const lines = String(markdownFile.markdown || "").split("\n");
                if (lines[0] === bookHeading) {
                    lines.shift();
                    while (lines[0] === "") {
                        lines.shift();
                    }
                }
                return lines.join("\n");
            }).filter((markdown) => markdown.trim().length > 0)
        ].join("\n\n");
        return { fileName, content };
    }
    function encodeMarkdownText(text, options = {}) {
        return textEncodingHelper.encodeText(text, options);
    }
    function createCombinedMarkdownExportPayload(workbook, markdownFiles, options = {}) {
        const combined = createCombinedMarkdownExportFile(workbook, markdownFiles);
        return {
            ...combined,
            data: encodeMarkdownText(`${combined.content}\n`, options),
            mimeType: textEncodingHelper.createTextMimeType(options)
        };
    }
    function createExportEntries(workbook, markdownFiles, options = {}) {
        const entries = [];
        if (markdownFiles.length > 0) {
            const combined = createCombinedMarkdownExportPayload(workbook, markdownFiles, options);
            entries.push({
                name: `output/${combined.fileName}`,
                data: combined.data
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
                if (!shape.svgPath || !shape.svgData)
                    continue;
                entries.push({
                    name: `output/${shape.svgPath}`,
                    data: shape.svgData
                });
            }
        }
        return entries;
    }
    function createWorkbookExportArchive(workbook, markdownFiles, options = {}) {
        return zipIoHelper.createStoredZip(createExportEntries(workbook, markdownFiles, options));
    }
    const markdownExportApi = {
        encodeMarkdownText,
        createCombinedMarkdownExportPayload,
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
