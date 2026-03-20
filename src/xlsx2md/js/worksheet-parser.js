(() => {
    const moduleRegistry = getXlsx2mdModuleRegistry();
    function extractCellOutputValue(cellElement, sharedStrings, cellStyle, deps, formulaOverride = "") {
        const type = (cellElement.getAttribute("t") || "").trim();
        const valueNode = cellElement.getElementsByTagName("v")[0] || null;
        const valueText = deps.getTextContent(valueNode);
        const formulaText = formulaOverride || deps.getTextContent(cellElement.getElementsByTagName("f")[0]);
        const cachedValueState = !formulaText
            ? null
            : !valueNode
                ? "absent"
                : valueText === ""
                    ? "present_empty"
                    : "present_nonempty";
        if (formulaText) {
            const normalizedFormula = formulaText.startsWith("=") ? formulaText : `=${formulaText}`;
            if (/\[[^\]]+\.xlsx\]/i.test(normalizedFormula)) {
                return {
                    valueType: type || "formula",
                    rawValue: valueText || normalizedFormula,
                    outputValue: normalizedFormula,
                    formulaText: normalizedFormula,
                    resolutionStatus: "unsupported_external",
                    resolutionSource: "external_unsupported",
                    cachedValueState
                };
            }
            if (valueNode) {
                const formattedValue = deps.formatCellDisplayValue(valueText, cellStyle);
                return {
                    valueType: type || "formula",
                    rawValue: valueText,
                    outputValue: formattedValue !== null && formattedValue !== void 0 ? formattedValue : valueText,
                    formulaText: normalizedFormula,
                    resolutionStatus: "resolved",
                    resolutionSource: "cached_value",
                    cachedValueState
                };
            }
            return {
                valueType: type || "formula",
                rawValue: normalizedFormula,
                outputValue: normalizedFormula,
                formulaText: normalizedFormula,
                resolutionStatus: "fallback_formula",
                resolutionSource: "formula_text",
                cachedValueState
            };
        }
        if (type === "s") {
            const sharedIndex = Number(valueText || 0);
            return {
                valueType: type,
                rawValue: valueText,
                outputValue: sharedStrings[sharedIndex] || "",
                formulaText: "",
                resolutionStatus: null,
                resolutionSource: null,
                cachedValueState: null
            };
        }
        if (type === "inlineStr") {
            const inlineText = Array.from(cellElement.getElementsByTagName("t")).map((node) => deps.getTextContent(node)).join("");
            return {
                valueType: type,
                rawValue: inlineText,
                outputValue: inlineText,
                formulaText: "",
                resolutionStatus: null,
                resolutionSource: null,
                cachedValueState: null
            };
        }
        if (type === "b") {
            return {
                valueType: type,
                rawValue: valueText,
                outputValue: valueText === "1" ? "TRUE" : "FALSE",
                formulaText: "",
                resolutionStatus: null,
                resolutionSource: null,
                cachedValueState: null
            };
        }
        if (type === "str" || type === "e") {
            return {
                valueType: type,
                rawValue: valueText,
                outputValue: valueText,
                formulaText: "",
                resolutionStatus: null,
                resolutionSource: null,
                cachedValueState: null
            };
        }
        if (valueText) {
            const formattedValue = deps.formatCellDisplayValue(valueText, cellStyle);
            if (formattedValue !== null) {
                return {
                    valueType: type,
                    rawValue: valueText,
                    outputValue: formattedValue,
                    formulaText: "",
                    resolutionStatus: null,
                    resolutionSource: null,
                    cachedValueState: null
                };
            }
        }
        return {
            valueType: type,
            rawValue: valueText,
            outputValue: valueText,
            formulaText: "",
            resolutionStatus: null,
            resolutionSource: null,
            cachedValueState: null
        };
    }
    function shiftReferenceAddress(addressText, rowOffset, colOffset, deps) {
        const match = String(addressText || "").match(/^(\$?)([A-Z]+)(\$?)(\d+)$/i);
        if (!match)
            return addressText;
        const colAbsolute = match[1] === "$";
        const rowAbsolute = match[3] === "$";
        const baseCol = deps.lettersToCol(match[2]);
        const baseRow = Number(match[4]);
        const shiftedCol = colAbsolute ? baseCol : baseCol + colOffset;
        const shiftedRow = rowAbsolute ? baseRow : baseRow + rowOffset;
        const safeCol = Math.max(1, shiftedCol);
        const safeRow = Math.max(1, shiftedRow);
        return `${colAbsolute ? "$" : ""}${deps.colToLetters(safeCol)}${rowAbsolute ? "$" : ""}${safeRow}`;
    }
    function translateSharedFormula(baseFormulaText, baseAddress, targetAddress, deps) {
        const basePos = deps.parseCellAddress(baseAddress);
        const targetPos = deps.parseCellAddress(targetAddress);
        if (!basePos.row || !basePos.col || !targetPos.row || !targetPos.col) {
            return baseFormulaText;
        }
        const rowOffset = targetPos.row - basePos.row;
        const colOffset = targetPos.col - basePos.col;
        const normalized = String(baseFormulaText || "").replace(/^=/, "");
        const translated = normalized.replace(/(?:'((?:[^']|'')+)'|([A-Za-z0-9_ ]+))!(\$?[A-Z]+\$?\d+)|(\$?[A-Z]+\$?\d+)/g, (full, quotedSheet, plainSheet, qualifiedAddress, localAddress) => {
            const address = qualifiedAddress || localAddress;
            if (!address)
                return full;
            const shifted = shiftReferenceAddress(address, rowOffset, colOffset, deps);
            if (qualifiedAddress) {
                const sheetPrefix = quotedSheet ? `'${quotedSheet}'` : plainSheet;
                return `${sheetPrefix}!${shifted}`;
            }
            return shifted;
        });
        return translated.startsWith("=") ? translated : `=${translated}`;
    }
    function parseWorksheet(files, sheetName, sheetPath, sheetIndex, sharedStrings, cellStyles, deps) {
        const bytes = files.get(sheetPath);
        if (!bytes) {
            throw new Error(`Sheet XML not found: ${sheetPath}`);
        }
        const doc = deps.xmlToDocument(deps.decodeXmlText(bytes));
        const sharedFormulaMap = new Map();
        const cells = Array.from(doc.getElementsByTagName("c")).map((cellElement) => {
            const address = cellElement.getAttribute("r") || "";
            const position = deps.parseCellAddress(address);
            const styleIndex = Number(cellElement.getAttribute("s") || 0);
            const cellStyle = cellStyles[styleIndex] || {
                borders: deps.EMPTY_BORDERS,
                numFmtId: 0,
                formatCode: "General"
            };
            let formulaOverride = "";
            const formulaElement = cellElement.getElementsByTagName("f")[0] || null;
            const formulaType = (formulaElement === null || formulaElement === void 0 ? void 0 : formulaElement.getAttribute("t")) || "";
            const spillRef = (formulaElement === null || formulaElement === void 0 ? void 0 : formulaElement.getAttribute("ref")) || "";
            const sharedIndex = (formulaElement === null || formulaElement === void 0 ? void 0 : formulaElement.getAttribute("si")) || "";
            const formulaText = deps.getTextContent(formulaElement);
            if (formulaType === "shared" && sharedIndex) {
                if (formulaText) {
                    const normalizedFormula = formulaText.startsWith("=") ? formulaText : `=${formulaText}`;
                    sharedFormulaMap.set(sharedIndex, { address, formulaText: normalizedFormula });
                    formulaOverride = normalizedFormula;
                }
                else {
                    const sharedBase = sharedFormulaMap.get(sharedIndex);
                    if (sharedBase) {
                        formulaOverride = translateSharedFormula(sharedBase.formulaText, sharedBase.address, address, deps);
                    }
                }
            }
            const output = extractCellOutputValue(cellElement, sharedStrings, cellStyle, deps, formulaOverride);
            return {
                address,
                row: position.row,
                col: position.col,
                valueType: output.valueType,
                rawValue: output.rawValue,
                outputValue: output.outputValue,
                formulaText: output.formulaText,
                resolutionStatus: output.resolutionStatus,
                resolutionSource: output.resolutionSource,
                cachedValueState: output.cachedValueState,
                styleIndex,
                borders: cellStyle.borders,
                numFmtId: cellStyle.numFmtId,
                formatCode: cellStyle.formatCode,
                formulaType,
                spillRef
            };
        });
        const merges = Array.from(doc.getElementsByTagName("mergeCell")).map((mergeElement) => deps.parseRangeRef(mergeElement.getAttribute("ref") || ""));
        const tables = deps.parseWorksheetTables(files, doc, sheetName, sheetPath);
        const assetDeps = deps.buildAssetDeps();
        const images = deps.parseDrawingImages(files, sheetName, sheetPath, assetDeps);
        const charts = deps.parseDrawingCharts(files, sheetName, sheetPath, assetDeps);
        const shapes = deps.parseDrawingShapes(files, sheetName, sheetPath, assetDeps);
        let maxRow = 0;
        let maxCol = 0;
        for (const cell of cells) {
            if (cell.row > maxRow)
                maxRow = cell.row;
            if (cell.col > maxCol)
                maxCol = cell.col;
        }
        for (const merge of merges) {
            if (merge.endRow > maxRow)
                maxRow = merge.endRow;
            if (merge.endCol > maxCol)
                maxCol = merge.endCol;
        }
        return {
            name: sheetName,
            index: sheetIndex,
            path: sheetPath,
            cells,
            merges,
            tables,
            images,
            charts,
            shapes,
            maxRow,
            maxCol
        };
    }
    const worksheetParserApi = {
        extractCellOutputValue,
        shiftReferenceAddress,
        translateSharedFormula,
        parseWorksheet
    };
    moduleRegistry.registerModule("worksheetParser", worksheetParserApi);
})();
