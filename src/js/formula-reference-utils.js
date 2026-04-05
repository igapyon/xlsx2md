/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const moduleRegistry = getXlsx2mdModuleRegistry();
    function createFormulaReferenceUtilsApi(deps) {
        function parseSimpleFormulaReference(formulaText, currentSheetName) {
            const normalizedFormula = String(formulaText || "").trim().replace(/^=/, "");
            const quotedSheetMatch = normalizedFormula.match(/^'((?:[^']|'')+)'!(\$?[A-Z]+\$?\d+)$/i);
            if (quotedSheetMatch) {
                return {
                    sheetName: quotedSheetMatch[1].replace(/''/g, "'"),
                    address: deps.normalizeFormulaAddress(quotedSheetMatch[2])
                };
            }
            const sheetMatch = normalizedFormula.match(/^([^'=][^!]*)!(\$?[A-Z]+\$?\d+)$/i);
            if (sheetMatch) {
                return {
                    sheetName: sheetMatch[1],
                    address: deps.normalizeFormulaAddress(sheetMatch[2])
                };
            }
            const localMatch = normalizedFormula.match(/^(\$?[A-Z]+\$?\d+)$/i);
            if (localMatch) {
                return {
                    sheetName: currentSheetName,
                    address: deps.normalizeFormulaAddress(localMatch[1])
                };
            }
            return null;
        }
        function normalizeFormulaSheetName(rawName) {
            return String(rawName || "").replace(/^'/, "").replace(/'$/, "").replace(/''/g, "'");
        }
        function normalizeDefinedNameKey(name) {
            return String(name || "").trim().toUpperCase();
        }
        function parseSheetScopedDefinedNameReference(expression, currentSheetName) {
            const normalizedExpression = String(expression || "").trim();
            const quotedSheetMatch = normalizedExpression.match(/^'((?:[^']|'')+)'!([A-Za-z_][A-Za-z0-9_.]*)$/);
            if (quotedSheetMatch) {
                return {
                    sheetName: normalizeFormulaSheetName(quotedSheetMatch[1].replace(/''/g, "'")),
                    name: quotedSheetMatch[2]
                };
            }
            const sheetMatch = normalizedExpression.match(/^([^'=][^!]*)!([A-Za-z_][A-Za-z0-9_.]*)$/);
            if (!sheetMatch) {
                return null;
            }
            if (/^\$?[A-Z]+\$?\d+$/i.test(sheetMatch[2])) {
                return null;
            }
            return {
                sheetName: normalizeFormulaSheetName(sheetMatch[1] || currentSheetName),
                name: sheetMatch[2]
            };
        }
        return {
            parseSimpleFormulaReference,
            parseSheetScopedDefinedNameReference,
            normalizeFormulaSheetName,
            normalizeDefinedNameKey
        };
    }
    const formulaReferenceUtilsApi = {
        createFormulaReferenceUtilsApi
    };
    moduleRegistry.registerModule("formulaReferenceUtils", formulaReferenceUtilsApi);
})();
