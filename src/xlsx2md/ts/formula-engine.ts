(() => {
  const moduleRegistry = getXlsx2mdModuleRegistry();
  type FormulaResolutionSource = "cached_value" | "ast_evaluator" | "legacy_resolver" | "formula_text" | "external_unsupported" | null;

  type FormulaEngineDeps = {
    getDefinedNameScalarValue: () => ((sheetName: string, name: string) => string | null) | null;
    tryResolveFormulaExpressionWithAst: (
      expression: string,
      currentSheetName: string,
      resolveCellValue: (sheetName: string, address: string) => string,
      resolveRangeEntries?: (sheetName: string, rangeText: string) => { rawValues: string[]; numericValues: number[] },
      currentAddress?: string
    ) => string | null;
    tryResolveFormulaExpressionLegacy: (
      normalized: string,
      currentSheetName: string,
      resolveCellValue: (sheetName: string, address: string) => string,
      resolveRangeValues?: (sheetName: string, rangeText: string) => number[],
      resolveRangeEntries?: (sheetName: string, rangeText: string) => { rawValues: string[]; numericValues: number[] }
    ) => string | null;
  };

  function createFormulaEngineApi(deps: FormulaEngineDeps) {
    function tryResolveFormulaExpressionDetailed(
      formulaText: string,
      currentSheetName: string,
      resolveCellValue: (sheetName: string, address: string) => string,
      resolveRangeValues?: (sheetName: string, rangeText: string) => number[],
      resolveRangeEntries?: (sheetName: string, rangeText: string) => { rawValues: string[]; numericValues: number[] },
      currentAddress?: string
    ): { value: string; source: FormulaResolutionSource } | null {
      const normalized = String(formulaText || "").trim().replace(/^=/, "");
      if (!normalized) return null;
      const directDefinedNameValue = deps.getDefinedNameScalarValue()?.(currentSheetName, normalized) || null;
      if (directDefinedNameValue != null) {
        return {
          value: directDefinedNameValue,
          source: "legacy_resolver"
        };
      }
      const astResolved = deps.tryResolveFormulaExpressionWithAst(
        normalized,
        currentSheetName,
        resolveCellValue,
        resolveRangeEntries,
        currentAddress
      );
      if (astResolved != null) {
        return {
          value: astResolved,
          source: "ast_evaluator"
        };
      }
      const legacyResolved = deps.tryResolveFormulaExpressionLegacy(
        normalized,
        currentSheetName,
        resolveCellValue,
        resolveRangeValues,
        resolveRangeEntries
      );
      if (legacyResolved == null) {
        return null;
      }
      return {
        value: legacyResolved,
        source: "legacy_resolver"
      };
    }

    function tryResolveFormulaExpression(
      formulaText: string,
      currentSheetName: string,
      resolveCellValue: (sheetName: string, address: string) => string,
      resolveRangeValues?: (sheetName: string, rangeText: string) => number[],
      resolveRangeEntries?: (sheetName: string, rangeText: string) => { rawValues: string[]; numericValues: number[] },
      currentAddress?: string
    ): string | null {
      return tryResolveFormulaExpressionDetailed(
        formulaText,
        currentSheetName,
        resolveCellValue,
        resolveRangeValues,
        resolveRangeEntries,
        currentAddress
      )?.value ?? null;
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
