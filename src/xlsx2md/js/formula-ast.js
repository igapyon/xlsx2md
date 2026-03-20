(() => {
    const moduleRegistry = getXlsx2mdModuleRegistry();
    function createFormulaAstApi(deps) {
        function tryResolveFormulaExpressionWithAst(expression, currentSheetName, resolveCellValue, resolveDefinedNameScalarValue, resolveDefinedNameRangeRef, resolveStructuredRangeRef, resolveSpillRange, resolveRangeEntries, currentAddress) {
            const formulaApi = moduleRegistry.getModule("formulaRuntime");
            if (!(formulaApi === null || formulaApi === void 0 ? void 0 : formulaApi.parseFormula) || !(formulaApi === null || formulaApi === void 0 ? void 0 : formulaApi.evaluateFormulaAst)) {
                return null;
            }
            try {
                const ast = formulaApi.parseFormula(`=${expression}`);
                const evaluated = formulaApi.evaluateFormulaAst(ast, {
                    resolveCell(ref, sheet) {
                        return coerceFormulaAstScalar(resolveCellValue(sheet || currentSheetName, deps.normalizeFormulaAddress(ref)));
                    },
                    resolveName(name) {
                        var _a, _b, _c;
                        const scopedRef = deps.parseSheetScopedDefinedNameReference(name, currentSheetName);
                        if (scopedRef) {
                            const scopedValue = (_a = resolveDefinedNameScalarValue === null || resolveDefinedNameScalarValue === void 0 ? void 0 : resolveDefinedNameScalarValue(scopedRef.sheetName, scopedRef.name)) !== null && _a !== void 0 ? _a : null;
                            if (scopedValue != null) {
                                return coerceFormulaAstScalar(scopedValue);
                            }
                        }
                        const scalarValue = (_b = resolveDefinedNameScalarValue === null || resolveDefinedNameScalarValue === void 0 ? void 0 : resolveDefinedNameScalarValue(currentSheetName, name)) !== null && _b !== void 0 ? _b : null;
                        if (scalarValue != null) {
                            return coerceFormulaAstScalar(scalarValue);
                        }
                        const rangeRef = (_c = resolveDefinedNameRangeRef === null || resolveDefinedNameRangeRef === void 0 ? void 0 : resolveDefinedNameRangeRef(currentSheetName, name)) !== null && _c !== void 0 ? _c : null;
                        if (rangeRef && resolveRangeEntries) {
                            return createFormulaAstRangeMatrix(rangeRef.sheetName, rangeRef.start, rangeRef.end, resolveRangeEntries);
                        }
                        return null;
                    },
                    resolveScopedName(sheet, name) {
                        var _a, _b;
                        const scopedValue = (_a = resolveDefinedNameScalarValue === null || resolveDefinedNameScalarValue === void 0 ? void 0 : resolveDefinedNameScalarValue(sheet, name)) !== null && _a !== void 0 ? _a : null;
                        if (scopedValue != null) {
                            return coerceFormulaAstScalar(scopedValue);
                        }
                        const rangeRef = (_b = resolveDefinedNameRangeRef === null || resolveDefinedNameRangeRef === void 0 ? void 0 : resolveDefinedNameRangeRef(sheet, name)) !== null && _b !== void 0 ? _b : null;
                        if (rangeRef && resolveRangeEntries) {
                            return createFormulaAstRangeMatrix(rangeRef.sheetName, rangeRef.start, rangeRef.end, resolveRangeEntries);
                        }
                        return null;
                    },
                    resolveStructuredRef(table, column) {
                        var _a;
                        const rangeRef = (_a = resolveStructuredRangeRef === null || resolveStructuredRangeRef === void 0 ? void 0 : resolveStructuredRangeRef(currentSheetName, `${table}[${column}]`)) !== null && _a !== void 0 ? _a : null;
                        if (!rangeRef || !resolveRangeEntries) {
                            return null;
                        }
                        return createFormulaAstRangeMatrix(rangeRef.sheetName, rangeRef.start, rangeRef.end, resolveRangeEntries);
                    },
                    resolveRange(startRef, endRef, sheet) {
                        if (!resolveRangeEntries) {
                            return [];
                        }
                        return createFormulaAstRangeMatrix(sheet || currentSheetName, deps.normalizeFormulaAddress(startRef), deps.normalizeFormulaAddress(endRef), resolveRangeEntries);
                    },
                    resolveSpill(ref, sheet) {
                        if (!resolveRangeEntries) {
                            return [];
                        }
                        const spillRange = resolveSpillRange(sheet || currentSheetName, ref);
                        if (!spillRange) {
                            return [];
                        }
                        return createFormulaAstRangeMatrix(spillRange.sheetName, spillRange.start, spillRange.end, resolveRangeEntries);
                    },
                    currentCellRef: currentAddress ? deps.normalizeFormulaAddress(currentAddress) : undefined
                });
                return serializeFormulaAstResult(evaluated);
            }
            catch (_error) {
                return null;
            }
        }
        function coerceFormulaAstScalar(value) {
            const trimmed = String(value || "").trim();
            if (!trimmed) {
                return "";
            }
            if (trimmed === "TRUE") {
                return true;
            }
            if (trimmed === "FALSE") {
                return false;
            }
            const numeric = Number(trimmed.replace(/,/g, ""));
            if (!Number.isNaN(numeric)) {
                return numeric;
            }
            return trimmed;
        }
        function createFormulaAstRangeMatrix(sheetName, startAddress, endAddress, resolveRangeEntries) {
            const range = deps.parseRangeAddress(`${deps.normalizeFormulaAddress(startAddress)}:${deps.normalizeFormulaAddress(endAddress)}`);
            if (!range) {
                return [];
            }
            const start = deps.parseCellAddress(range.start);
            const end = deps.parseCellAddress(range.end);
            if (!start.row || !start.col || !end.row || !end.col) {
                return [];
            }
            const startRow = Math.min(start.row, end.row);
            const endRow = Math.max(start.row, end.row);
            const startCol = Math.min(start.col, end.col);
            const endCol = Math.max(start.col, end.col);
            const entries = resolveRangeEntries(sheetName, `${range.start}:${range.end}`).rawValues;
            const matrix = [];
            let index = 0;
            for (let row = startRow; row <= endRow; row += 1) {
                const rowValues = [];
                for (let col = startCol; col <= endCol; col += 1) {
                    rowValues.push(coerceFormulaAstScalar(entries[index] || ""));
                    index += 1;
                }
                matrix.push(rowValues);
            }
            return matrix;
        }
        function serializeFormulaAstResult(value) {
            if (value == null) {
                return null;
            }
            if (Array.isArray(value)) {
                return null;
            }
            if (typeof value === "boolean") {
                return value ? "TRUE" : "FALSE";
            }
            return String(value);
        }
        return {
            tryResolveFormulaExpressionWithAst,
            coerceFormulaAstScalar,
            createFormulaAstRangeMatrix,
            serializeFormulaAstResult
        };
    }
    const formulaAstApi = {
        createFormulaAstApi
    };
    moduleRegistry.registerModule("formulaAst", formulaAstApi);
})();
