/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const moduleRegistry = getXlsx2mdModuleRegistry();
    function createFormulaEngineApi(deps) {
        function tryResolveFormulaExpressionDetailed(formulaText, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries, currentAddress) {
            var _a;
            const normalized = String(formulaText || "").trim().replace(/^=/, "");
            if (!normalized)
                return null;
            const directDefinedNameValue = ((_a = deps.getDefinedNameScalarValue()) === null || _a === void 0 ? void 0 : _a(currentSheetName, normalized)) || null;
            if (directDefinedNameValue != null) {
                return {
                    value: directDefinedNameValue,
                    source: "legacy_resolver"
                };
            }
            const astResolved = deps.tryResolveFormulaExpressionWithAst(normalized, currentSheetName, resolveCellValue, resolveRangeEntries, currentAddress);
            if (astResolved != null) {
                return {
                    value: astResolved,
                    source: "ast_evaluator"
                };
            }
            const legacyResolved = deps.tryResolveFormulaExpressionLegacy(normalized, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
            if (legacyResolved == null) {
                return null;
            }
            return {
                value: legacyResolved,
                source: "legacy_resolver"
            };
        }
        function tryResolveFormulaExpression(formulaText, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries, currentAddress) {
            var _a, _b;
            return (_b = (_a = tryResolveFormulaExpressionDetailed(formulaText, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries, currentAddress)) === null || _a === void 0 ? void 0 : _a.value) !== null && _b !== void 0 ? _b : null;
        }
        return {
            tryResolveFormulaExpressionDetailed,
            tryResolveFormulaExpression
        };
    }
    const formulaEngineApi = {
        createFormulaEngineApi
    };
    moduleRegistry.registerModule("formulaEngine", formulaEngineApi);
})();
