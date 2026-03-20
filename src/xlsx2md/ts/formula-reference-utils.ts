(() => {
  const moduleRegistry = getXlsx2mdModuleRegistry();
  type FormulaReferenceUtilsDeps = {
    normalizeFormulaAddress: (address: string) => string;
  };

  function createFormulaReferenceUtilsApi(deps: FormulaReferenceUtilsDeps) {
    function parseSimpleFormulaReference(
      formulaText: string,
      currentSheetName: string
    ): { sheetName: string; address: string } | null {
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

    function normalizeFormulaSheetName(rawName: string): string {
      return String(rawName || "").replace(/^'/, "").replace(/'$/, "").replace(/''/g, "'");
    }

    function normalizeDefinedNameKey(name: string): string {
      return String(name || "").trim().toUpperCase();
    }

    function parseSheetScopedDefinedNameReference(
      expression: string,
      currentSheetName: string
    ): { sheetName: string; name: string } | null {
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
