(() => {
  const moduleRegistry = getXlsx2mdModuleRegistry();
  type FormulaResolutionSource = "cached_value" | "ast_evaluator" | "legacy_resolver" | "formula_text" | "external_unsupported" | null;

  type FormulaAstDeps = {
    normalizeFormulaAddress: (address: string) => string;
    parseSheetScopedDefinedNameReference: (
      formulaText: string,
      currentSheetName: string
    ) => { sheetName: string; name: string } | null;
    parseRangeAddress: (rawRange: string) => { start: string; end: string } | null;
    parseCellAddress: (address: string) => { row: number; col: number };
  };

  function createFormulaAstApi(deps: FormulaAstDeps) {
    function tryResolveFormulaExpressionWithAst(
      expression: string,
      currentSheetName: string,
      resolveCellValue: (sheetName: string, address: string) => string,
      resolveDefinedNameScalarValue: ((sheetName: string, name: string) => string | null) | null,
      resolveDefinedNameRangeRef: ((sheetName: string, name: string) => { sheetName: string; start: string; end: string } | null) | null,
      resolveStructuredRangeRef: ((sheetName: string, text: string) => { sheetName: string; start: string; end: string } | null) | null,
      resolveSpillRange: ((sheetName: string, ref: string) => { sheetName: string; start: string; end: string } | null),
      resolveRangeEntries?: (sheetName: string, rangeText: string) => { rawValues: string[]; numericValues: number[] },
      currentAddress?: string
    ): string | null {
      const formulaApi = moduleRegistry.getModule<Record<string, unknown>>("formulaRuntime");
      if (!formulaApi?.parseFormula || !formulaApi?.evaluateFormulaAst) {
        return null;
      }
      try {
        const ast = formulaApi.parseFormula(`=${expression}`);
        const evaluated = formulaApi.evaluateFormulaAst(ast, {
          resolveCell(ref: string, sheet: string | null) {
            return coerceFormulaAstScalar(resolveCellValue(sheet || currentSheetName, deps.normalizeFormulaAddress(ref)));
          },
          resolveName(name: string) {
            const scopedRef = deps.parseSheetScopedDefinedNameReference(name, currentSheetName);
            if (scopedRef) {
              const scopedValue = resolveDefinedNameScalarValue?.(scopedRef.sheetName, scopedRef.name) ?? null;
              if (scopedValue != null) {
                return coerceFormulaAstScalar(scopedValue);
              }
            }
            const scalarValue = resolveDefinedNameScalarValue?.(currentSheetName, name) ?? null;
            if (scalarValue != null) {
              return coerceFormulaAstScalar(scalarValue);
            }
            const rangeRef = resolveDefinedNameRangeRef?.(currentSheetName, name) ?? null;
            if (rangeRef && resolveRangeEntries) {
              return createFormulaAstRangeMatrix(
                rangeRef.sheetName,
                rangeRef.start,
                rangeRef.end,
                resolveRangeEntries
              );
            }
            return null;
          },
          resolveScopedName(sheet: string, name: string) {
            const scopedValue = resolveDefinedNameScalarValue?.(sheet, name) ?? null;
            if (scopedValue != null) {
              return coerceFormulaAstScalar(scopedValue);
            }
            const rangeRef = resolveDefinedNameRangeRef?.(sheet, name) ?? null;
            if (rangeRef && resolveRangeEntries) {
              return createFormulaAstRangeMatrix(
                rangeRef.sheetName,
                rangeRef.start,
                rangeRef.end,
                resolveRangeEntries
              );
            }
            return null;
          },
          resolveStructuredRef(table: string, column: string) {
            const rangeRef = resolveStructuredRangeRef?.(currentSheetName, `${table}[${column}]`) ?? null;
            if (!rangeRef || !resolveRangeEntries) {
              return null;
            }
            return createFormulaAstRangeMatrix(
              rangeRef.sheetName,
              rangeRef.start,
              rangeRef.end,
              resolveRangeEntries
            );
          },
          resolveRange(startRef: string, endRef: string, sheet: string | null) {
            if (!resolveRangeEntries) {
              return [];
            }
            return createFormulaAstRangeMatrix(
              sheet || currentSheetName,
              deps.normalizeFormulaAddress(startRef),
              deps.normalizeFormulaAddress(endRef),
              resolveRangeEntries
            );
          },
          resolveSpill(ref: string, sheet: string | null) {
            if (!resolveRangeEntries) {
              return [];
            }
            const spillRange = resolveSpillRange(sheet || currentSheetName, ref);
            if (!spillRange) {
              return [];
            }
            return createFormulaAstRangeMatrix(
              spillRange.sheetName,
              spillRange.start,
              spillRange.end,
              resolveRangeEntries
            );
          },
          currentCellRef: currentAddress ? deps.normalizeFormulaAddress(currentAddress) : undefined
        });
        return serializeFormulaAstResult(evaluated);
      } catch (_error) {
        return null;
      }
    }

    function coerceFormulaAstScalar(value: string): string | number | boolean {
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

    function createFormulaAstRangeMatrix(
      sheetName: string,
      startAddress: string,
      endAddress: string,
      resolveRangeEntries: (sheetName: string, rangeText: string) => { rawValues: string[]; numericValues: number[] }
    ): (string | number | boolean)[][] {
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
      const matrix: (string | number | boolean)[][] = [];
      let index = 0;
      for (let row = startRow; row <= endRow; row += 1) {
        const rowValues: (string | number | boolean)[] = [];
        for (let col = startCol; col <= endCol; col += 1) {
          rowValues.push(coerceFormulaAstScalar(entries[index] || ""));
          index += 1;
        }
        matrix.push(rowValues);
      }
      return matrix;
    }

    function serializeFormulaAstResult(value: unknown): string | null {
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
