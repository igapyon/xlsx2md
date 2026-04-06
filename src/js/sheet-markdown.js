/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const moduleRegistry = getXlsx2mdModuleRegistry();
    function createSheetMarkdownApi(deps) {
        const richTextRenderer = requireXlsx2mdRichTextRendererModule().createRichTextRendererApi({
            normalizeMarkdownText: deps.normalizeMarkdownText
        });
        function buildCellMap(sheet) {
            const map = new Map();
            for (const cell of sheet.cells) {
                map.set(`${cell.row}:${cell.col}`, cell);
            }
            return map;
        }
        function createHeadingFragment(text) {
            return String(text || "")
                .trim()
                .toLowerCase()
                .replace(/<[^>]+>/g, "")
                .replace(/[^\p{L}\p{N}\s_-]+/gu, "")
                .replace(/\s+/g, "-");
        }
        function parseInternalHyperlinkLocation(location, currentSheetName) {
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
        function renderHyperlinkMarkdown(cell, text, workbook, sheet, options) {
            const hyperlink = cell.hyperlink;
            const label = String(text || "").trim();
            if (!hyperlink || !label)
                return text;
            if (hyperlink.kind === "external") {
                const href = String(hyperlink.target || "").trim();
                return href ? `[${label}](${href})` : label;
            }
            const currentSheetName = (sheet === null || sheet === void 0 ? void 0 : sheet.name) || "";
            const { sheetName, refText } = parseInternalHyperlinkLocation(hyperlink.location || hyperlink.target, currentSheetName);
            const traceText = [sheetName, refText].filter(Boolean).join("!");
            const targetSheet = (workbook === null || workbook === void 0 ? void 0 : workbook.sheets.find((entry) => entry.name === sheetName)) || null;
            if (!targetSheet || !workbook) {
                return traceText ? `${label} (${traceText})` : label;
            }
            const href = `#${createHeadingFragment(targetSheet.name)}`;
            return traceText && traceText !== targetSheet.name
                ? `[${label}](${href}) (${traceText})`
                : `[${label}](${href})`;
        }
        function formatCellForMarkdown(cell, options, workbook = null, sheet = null) {
            var _a;
            if (!cell)
                return "";
            const mode = options.outputMode || "display";
            const formattingMode = options.formattingMode || "plain";
            const displayCell = formattingMode === "github" && cell.hyperlink
                ? {
                    ...cell,
                    textStyle: {
                        ...cell.textStyle,
                        underline: false
                    },
                    richTextRuns: ((_a = cell.richTextRuns) === null || _a === void 0 ? void 0 : _a.map((run) => ({
                        ...run,
                        underline: false
                    }))) || null
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
        function isCellInAnyTable(row, col, tables) {
            return tables.some((table) => row >= table.startRow && row <= table.endRow && col >= table.startCol && col <= table.endCol);
        }
        function splitNarrativeRowSegments(cells, options, workbook = null, sheet = null) {
            const segments = [];
            let current = null;
            for (const cell of cells) {
                const value = formatCellForMarkdown(cell, options, workbook, sheet).trim();
                if (!value)
                    continue;
                if (!current || cell.col - current.lastCol > 4) {
                    current = {
                        startCol: cell.col,
                        values: [value],
                        lastCol: cell.col
                    };
                    segments.push(current);
                }
                else {
                    current.values.push(value);
                    current.lastCol = cell.col;
                }
            }
            return segments.map((segment) => ({
                startCol: segment.startCol,
                values: segment.values
            }));
        }
        function extractNarrativeBlocks(workbook, sheet, tables, options = {}) {
            const rowMap = new Map();
            for (const cell of sheet.cells) {
                if (!cell.outputValue)
                    continue;
                if (isCellInAnyTable(cell.row, cell.col, tables))
                    continue;
                const entries = rowMap.get(cell.row) || [];
                entries.push(cell);
                rowMap.set(cell.row, entries);
            }
            const rowNumbers = Array.from(rowMap.keys()).sort((a, b) => a - b);
            const blocks = [];
            let current = null;
            let previousRow = -100;
            for (const rowNumber of rowNumbers) {
                const cells = (rowMap.get(rowNumber) || []).slice().sort((a, b) => a.col - b.col);
                const rowSegments = splitNarrativeRowSegments(cells, options, workbook, sheet);
                for (const segment of rowSegments) {
                    const rowText = segment.values.join(" ").trim();
                    if (!rowText)
                        continue;
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
                    }
                    else {
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
        function extractSectionBlocks(sheet, tables, narrativeBlocks) {
            const charts = sheet.charts || [];
            const anchors = [];
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
                if (left.startRow !== right.startRow)
                    return left.startRow - right.startRow;
                return left.startCol - right.startCol;
            });
            const sections = [];
            let current = null;
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
                }
                else {
                    current.startRow = Math.min(current.startRow, anchor.startRow);
                    current.startCol = Math.min(current.startCol, anchor.startCol);
                    current.endRow = Math.max(current.endRow, anchor.endRow);
                    current.endCol = Math.max(current.endCol, anchor.endCol);
                }
                previousEndRow = Math.max(previousEndRow, anchor.endRow);
            }
            return sections;
        }
        function convertSheetToMarkdown(workbook, sheet, options = {}) {
            var _a;
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
            const sections = [];
            for (const block of narrativeBlocks) {
                sections.push({
                    sortRow: block.startRow,
                    sortCol: block.startCol,
                    markdown: `${deps.renderNarrativeBlock(block)}\n`,
                    kind: "narrative",
                    narrativeBlock: block
                });
            }
            const fileName = deps.createOutputFileName(workbook.name, sheet.index, sheet.name, options.outputMode || "display", options.formattingMode || "plain");
            let tableCounter = 1;
            for (const table of tables) {
                const rows = deps.matrixFromCandidate(sheet, table, options, buildCellMap, (cell, tableOptions) => formatCellForMarkdown(cell, tableOptions, workbook, sheet));
                if (rows.length === 0 || ((_a = rows[0]) === null || _a === void 0 ? void 0 : _a.length) === 0)
                    continue;
                const tableMarkdown = deps.renderMarkdownTable(rows, treatFirstRowAsHeader);
                sections.push({
                    sortRow: table.startRow,
                    sortCol: table.startCol,
                    markdown: `### Table: ${String(tableCounter).padStart(3, "0")} (${deps.formatRange(table.startRow, table.startCol, table.endRow, table.endCol)})\n\n${tableMarkdown}\n`,
                    kind: "table"
                });
                tableCounter += 1;
            }
            sections.sort((left, right) => {
                if (left.sortRow !== right.sortRow)
                    return left.sortRow - right.sortRow;
                return left.sortCol - right.sortCol;
            });
            const groupedSections = (sectionBlocks.length > 0 ? sectionBlocks : [{
                    startRow: -1,
                    startCol: -1,
                    endRow: Number.MAX_SAFE_INTEGER,
                    endCol: Number.MAX_SAFE_INTEGER
                }]).map((block) => ({
                block,
                entries: sections.filter((section) => section.sortRow >= block.startRow
                    && section.sortRow <= block.endRow
                    && section.sortCol >= block.startCol
                    && section.sortCol <= block.endCol)
            })).filter((group) => group.entries.length > 0);
            const body = groupedSections
                .map((group) => group.entries.map((section) => section.markdown.trimEnd()).join("\n\n").trim())
                .filter(Boolean)
                .join("\n\n---\n\n")
                .trim();
            const imageSection = sheet.images.length > 0
                ? [
                    "",
                    ...sheet.images.map((image, index) => [
                        `### Image: ${String(index + 1).padStart(3, "0")} (${image.anchor})`,
                        `- File: ${image.path}`,
                        "",
                        `![${image.filename}](${image.path})`
                    ].join("\n"))
                ].join("\n\n")
                : "";
            const chartSection = charts.length > 0
                ? [
                    "",
                    ...charts.map((chart, index) => {
                        const lines = [
                            `### Chart: ${String(index + 1).padStart(3, "0")} (${chart.anchor})`,
                            `- Title: ${chart.title || "(none)"}`,
                            `- Type: ${chart.chartType}`
                        ];
                        if (chart.series.length > 0) {
                            lines.push("- Series:");
                            for (const series of chart.series) {
                                lines.push(`  - ${series.name}`);
                                if (series.axis === "secondary")
                                    lines.push("    - Axis: secondary");
                                if (series.categoriesRef)
                                    lines.push(`    - categories: ${series.categoriesRef}`);
                                if (series.valuesRef)
                                    lines.push(`    - values: ${series.valuesRef}`);
                            }
                        }
                        return lines.join("\n");
                    })
                ].join("\n\n")
                : "";
            const includeShapeDetails = options.includeShapeDetails !== false;
            const shapeSection = includeShapeDetails && shapes.length > 0
                ? [
                    "",
                    ...shapeBlocks.map((block, blockIndex) => {
                        const shapeDetails = block.shapeIndexes
                            .map((shapeIndex) => {
                            const shape = shapes[shapeIndex];
                            if (!shape)
                                return "";
                            const lines = [
                                `#### Shape: ${String(shapeIndex + 1).padStart(3, "0")} (${shape.anchor})`,
                                ...deps.renderHierarchicalRawEntries(shape.rawEntries)
                            ];
                            if (shape.svgPath) {
                                lines.push(`- SVG: ${shape.svgPath}`);
                                lines.push("");
                                lines.push(`![${shape.svgFilename || `shape_${String(shapeIndex + 1).padStart(3, "0")}.svg`}](${shape.svgPath})`);
                            }
                            return lines.join("\n");
                        })
                            .filter(Boolean)
                            .join("\n\n");
                        return [
                            `### Shape Block: ${String(blockIndex + 1).padStart(3, "0")} (${deps.formatRange(block.startRow, block.startCol, block.endRow, block.endCol)})`,
                            `- Shapes: ${block.shapeIndexes.map((shapeIndex) => `Shape ${String(shapeIndex + 1).padStart(3, "0")}`).join(", ")}`,
                            `- anchorRange: ${deps.colToLetters(block.startCol)}${block.startRow}-${deps.colToLetters(block.endCol)}${block.endRow}`,
                            ...(shapeDetails ? ["", shapeDetails] : [])
                        ].join("\n");
                    }),
                    ...(() => {
                        const grouped = new Set(shapeBlocks.flatMap((block) => block.shapeIndexes));
                        const ungrouped = shapes
                            .map((shape, index) => ({ shape, index }))
                            .filter(({ index }) => !grouped.has(index));
                        if (ungrouped.length === 0) {
                            return [];
                        }
                        return [
                            "",
                            "### Ungrouped Shapes",
                            "",
                            ...ungrouped.map(({ shape, index }) => {
                                const lines = [
                                    `#### Shape: ${String(index + 1).padStart(3, "0")} (${shape.anchor})`,
                                    ...deps.renderHierarchicalRawEntries(shape.rawEntries)
                                ];
                                if (shape.svgPath) {
                                    lines.push(`- SVG: ${shape.svgPath}`);
                                    lines.push("");
                                    lines.push(`![${shape.svgFilename || `shape_${String(index + 1).padStart(3, "0")}.svg`}](${shape.svgPath})`);
                                }
                                return lines.join("\n");
                            })
                        ];
                    })()
                ].join("\n\n")
                : "";
            const markdown = [
                `# Book: ${workbook.name}`,
                "",
                `## Sheet: ${sheet.name}`,
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
        function convertWorkbookToMarkdownFiles(workbook, options = {}) {
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
