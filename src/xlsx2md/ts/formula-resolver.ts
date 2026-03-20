(() => {
  const moduleRegistry = getXlsx2mdModuleRegistry();
  type FormulaResolutionStatus = "resolved" | "fallback_formula" | "unsupported_external" | null;
  type FormulaResolutionSource = "cached_value" | "ast_evaluator" | "legacy_resolver" | "formula_text" | "external_unsupported" | null;

  type ParsedCell = {
    address: string;
    rawValue: string;
    outputValue: string;
    formulaText: string;
    resolutionStatus: FormulaResolutionStatus;
    resolutionSource: FormulaResolutionSource;
    valueType: string;
    formulaType: string;
    spillRef: string;
  };

  type ParsedTable = {
    sheetName: string;
    name: string;
    displayName: string;
    start: string;
    end: string;
    columns: string[];
    headerRowCount: number;
    totalsRowCount: number;
  };

  type ParsedSheet = {
    name: string;
    cells: ParsedCell[];
    tables: ParsedTable[];
  };

  type ParsedWorkbook = {
    sheets: ParsedSheet[];
    definedNames: {
      name: string;
      formulaText: string;
      localSheetName: string | null;
    }[];
  };

  type ResolverDeps = {
    normalizeStructuredTableKey: (value: string) => string;
    normalizeFormulaSheetName: (rawName: string) => string;
    normalizeDefinedNameKey: (name: string) => string;
    normalizeFormulaAddress: (address: string) => string;
    parseSimpleFormulaReference: (formulaText: string, currentSheetName: string) => { sheetName: string; address: string } | null;
    resolveScalarFormulaValue: (
      expression: string,
      currentSheetName: string,
      resolveCellValue: (sheetName: string, address: string) => string,
      resolveRangeValues?: (sheetName: string, rangeText: string) => number[],
      resolveRangeEntries?: (sheetName: string, rangeText: string) => { rawValues: string[]; numericValues: number[] }
    ) => string | null;
    parseQualifiedRangeReference: (text: string, currentSheetName: string) => { sheetName: string; start: string; end: string } | null;
    findTopLevelOperatorIndex: (expression: string, operator: string) => number;
    parseWholeFunctionCall: (formulaText: string, targetNames: string[]) => { name: string; argsText: string } | null;
    splitFormulaArguments: (argText: string) => string[];
    parseCellAddress: (address: string) => { row: number; col: number };
    colToLetters: (col: number) => string;
    parseRangeAddress: (rawRange: string) => { start: string; end: string } | null;
    tryResolveFormulaExpressionDetailed: (
      formulaText: string,
      currentSheetName: string,
      resolveCellValue: (sheetName: string, address: string) => string,
      resolveRangeValues?: (sheetName: string, rangeText: string) => number[],
      resolveRangeEntries?: (sheetName: string, rangeText: string) => { rawValues: string[]; numericValues: number[] },
      currentAddress?: string
    ) => { value: string; source: FormulaResolutionSource } | null;
    applyResolvedFormulaValue: (cell: ParsedCell, resolvedValue: string, resolutionSource?: FormulaResolutionSource) => void;
    setDefinedNameResolvers?: (
      scalar: ((sheetName: string, name: string) => string | null) | null,
      range: ((sheetName: string, name: string) => { sheetName: string; start: string; end: string } | null) | null,
      structured: ((sheetName: string, text: string) => { sheetName: string; start: string; end: string } | null) | null
    ) => void;
  };

  function buildFormulaResolver(workbook: ParsedWorkbook, deps: ResolverDeps) {
    const sheetMap = new Map<string, ParsedSheet>();
    const cellMaps = new Map<string, Map<string, ParsedCell>>();
    const tableMap = new Map<string, ParsedTable>();
    for (const sheet of workbook.sheets) {
      sheetMap.set(sheet.name, sheet);
      const cellMap = new Map<string, ParsedCell>();
      for (const cell of sheet.cells) {
        cellMap.set(cell.address.toUpperCase(), cell);
      }
      cellMaps.set(sheet.name, cellMap);
      for (const table of sheet.tables) {
        if (table.name) {
          tableMap.set(deps.normalizeStructuredTableKey(table.name), table);
        }
        if (table.displayName) {
          tableMap.set(deps.normalizeStructuredTableKey(table.displayName), table);
        }
      }
    }

    const resolvingKeys = new Set<string>();
    const definedNameMap = new Map<string, string>();
    for (const entry of workbook.definedNames) {
      const key = entry.localSheetName
        ? `${deps.normalizeFormulaSheetName(entry.localSheetName)}::${deps.normalizeDefinedNameKey(entry.name)}`
        : `::${deps.normalizeDefinedNameKey(entry.name)}`;
      definedNameMap.set(key, entry.formulaText);
    }

    function lookupDefinedNameFormula(sheetName: string, name: string): string | null {
      const normalizedName = deps.normalizeDefinedNameKey(name);
      return definedNameMap.get(`${deps.normalizeFormulaSheetName(sheetName)}::${normalizedName}`)
        || definedNameMap.get(`::${normalizedName}`)
        || null;
    }

    function resolveCellValue(sheetName: string, address: string): string {
      const sheet = sheetMap.get(sheetName);
      if (!sheet) return "#REF!";
      const cell = cellMaps.get(sheetName)?.get(address.toUpperCase()) || null;
      if (!cell) return "";
      const key = `${sheetName}!${address.toUpperCase()}`;
      if (resolvingKeys.has(key)) {
        return "";
      }
      if (cell.formulaText && (!cell.outputValue || cell.resolutionStatus !== "resolved")) {
        resolvingKeys.add(key);
        try {
          try {
            const result = deps.tryResolveFormulaExpressionDetailed(
              cell.formulaText,
              sheetName,
              resolveCellValue,
              undefined,
              undefined,
              cell.address
            );
            if (result?.value != null) {
              deps.applyResolvedFormulaValue(cell, result.value, result.source || "legacy_resolver");
            }
          } catch (error) {
            if (!(error instanceof Error) || error.message !== "__FORMULA_UNRESOLVED__") {
              throw error;
            }
          }
        } finally {
          resolvingKeys.delete(key);
        }
      }
      if (cell.formulaText) {
        if (cell.resolutionStatus === "resolved") {
          const rawValue = String(cell.rawValue || "");
          const outputValue = String(cell.outputValue || "");
          if (rawValue && rawValue !== cell.formulaText) {
            return rawValue;
          }
          return outputValue || rawValue;
        }
        const rawValue = String(cell.rawValue || "");
        const outputValue = String(cell.outputValue || "");
        if (rawValue && rawValue !== cell.formulaText) {
          return rawValue;
        }
        if (outputValue && outputValue !== cell.formulaText) {
          return outputValue;
        }
        return "";
      }
      if (["s", "inlineStr", "str", "e", "b"].includes(cell.valueType)) {
        return String(cell.outputValue || cell.rawValue || "");
      }
      return String(cell.rawValue || cell.outputValue || "");
    }

    function resolveRangeEntries(sheetName: string, rangeText: string): { rawValues: string[]; numericValues: number[] } {
      const range = deps.parseRangeAddress(rangeText);
      if (!range) {
        return { rawValues: [], numericValues: [] };
      }
      const start = deps.parseCellAddress(range.start);
      const end = deps.parseCellAddress(range.end);
      if (!start.row || !start.col || !end.row || !end.col) {
        return { rawValues: [], numericValues: [] };
      }
      const startRow = Math.min(start.row, end.row);
      const endRow = Math.max(start.row, end.row);
      const startCol = Math.min(start.col, end.col);
      const endCol = Math.max(start.col, end.col);
      const rawValues: string[] = [];
      const numericValues: number[] = [];
      for (let row = startRow; row <= endRow; row += 1) {
        for (let col = startCol; col <= endCol; col += 1) {
          const rawValue = resolveCellValue(sheetName, `${deps.colToLetters(col)}${row}`);
          rawValues.push(rawValue);
          if (String(rawValue || "").trim() === "") continue;
          const numericValue = Number(rawValue);
          if (!Number.isNaN(numericValue)) {
            numericValues.push(numericValue);
          }
        }
      }
      return { rawValues, numericValues };
    }

    function resolveDefinedNameValue(sheetName: string, name: string): string | null {
      const formulaText = lookupDefinedNameFormula(sheetName, name);
      if (!formulaText) return null;
      const directRef = deps.parseSimpleFormulaReference(formulaText, sheetName);
      if (directRef) {
        const value = resolveCellValue(directRef.sheetName, directRef.address);
        return value === "" ? null : value;
      }
      const scalar = deps.resolveScalarFormulaValue(formulaText.replace(/^=/, ""), sheetName, resolveCellValue);
      return scalar == null || scalar === "" ? null : scalar;
    }

    function resolveDefinedNameRange(sheetName: string, name: string): { sheetName: string; start: string; end: string } | null {
      const formulaText = lookupDefinedNameFormula(sheetName, name);
      if (!formulaText) return null;
      const normalized = formulaText.replace(/^=/, "").trim();
      const directRange = deps.parseQualifiedRangeReference(normalized, sheetName);
      if (directRange) {
        return directRange;
      }
      const separatorIndex = deps.findTopLevelOperatorIndex(normalized, ":");
      if (separatorIndex <= 0) return null;
      const leftText = normalized.slice(0, separatorIndex).trim();
      const rightText = normalized.slice(separatorIndex + 1).trim();
      const startRef = deps.parseSimpleFormulaReference(`=${leftText}`, sheetName);
      const indexCall = deps.parseWholeFunctionCall(rightText, ["INDEX"]);
      if (!startRef || !indexCall) return null;
      const args = deps.splitFormulaArguments(indexCall.argsText.trim());
      if (args.length < 2 || args.length > 3) return null;
      const rangeRef = deps.parseQualifiedRangeReference(args[0], sheetName);
      const rowIndex = Number(deps.resolveScalarFormulaValue(
        args[1],
        sheetName,
        resolveCellValue,
        (targetSheetName, rangeText) => resolveRangeEntries(targetSheetName, rangeText).numericValues,
        resolveRangeEntries
      ));
      const colIndex = args.length === 3
        ? Number(deps.resolveScalarFormulaValue(
          args[2],
          sheetName,
          resolveCellValue,
          (targetSheetName, rangeText) => resolveRangeEntries(targetSheetName, rangeText).numericValues,
          resolveRangeEntries
        ))
        : 1;
      if (!rangeRef || Number.isNaN(rowIndex) || Number.isNaN(colIndex) || rowIndex < 1 || colIndex < 1) return null;
      const rangeStart = deps.parseCellAddress(rangeRef.start);
      const rangeEnd = deps.parseCellAddress(rangeRef.end);
      if (!rangeStart.row || !rangeStart.col || !rangeEnd.row || !rangeEnd.col) return null;
      const startRow = Math.min(rangeStart.row, rangeEnd.row);
      const endRow = Math.max(rangeStart.row, rangeEnd.row);
      const startCol = Math.min(rangeStart.col, rangeEnd.col);
      const endCol = Math.max(rangeStart.col, rangeEnd.col);
      const targetRow = startRow + Math.trunc(rowIndex) - 1;
      const targetCol = startCol + Math.trunc(colIndex) - 1;
      if (targetRow > endRow || targetCol > endCol) return null;
      return {
        sheetName: startRef.sheetName,
        start: startRef.address,
        end: `${deps.colToLetters(targetCol)}${targetRow}`
      };
    }

    function resolveStructuredRange(sheetName: string, text: string): { sheetName: string; start: string; end: string } | null {
      const match = String(text || "").trim().match(/^(.+?)\[([^\]]+)\]$/);
      if (!match) return null;
      const tableKey = deps.normalizeStructuredTableKey(match[1].replace(/^'(.*)'$/, "$1"));
      const columnKey = deps.normalizeStructuredTableKey(match[2]);
      if (!tableKey || !columnKey || columnKey.startsWith("#") || columnKey.startsWith("@")) return null;
      const table = tableMap.get(tableKey);
      if (!table) return null;
      const columnIndex = table.columns.findIndex((columnName) => deps.normalizeStructuredTableKey(columnName) === columnKey);
      if (columnIndex < 0) return null;
      const startAddress = deps.parseCellAddress(table.start);
      const endAddress = deps.parseCellAddress(table.end);
      if (!startAddress.row || !startAddress.col || !endAddress.row || !endAddress.col) return null;
      const firstDataRow = Math.min(startAddress.row, endAddress.row) + Math.max(0, table.headerRowCount);
      const lastDataRow = Math.max(startAddress.row, endAddress.row) - Math.max(0, table.totalsRowCount);
      if (firstDataRow > lastDataRow) return null;
      const col = Math.min(startAddress.col, endAddress.col) + columnIndex;
      const colLetters = deps.colToLetters(col);
      return {
        sheetName: table.sheetName || sheetName,
        start: `${colLetters}${firstDataRow}`,
        end: `${colLetters}${lastDataRow}`
      };
    }

    return {
      resolveCellValue,
      resolveRangeValues: (sheetName: string, rangeText: string) => resolveRangeEntries(sheetName, rangeText).numericValues,
      resolveRangeEntries,
      resolveDefinedNameValue,
      resolveDefinedNameRange,
      resolveStructuredRange
    };
  }

  function resolveSimpleFormulaReferences(workbook: ParsedWorkbook, deps: ResolverDeps): void {
    const resolver = buildFormulaResolver(workbook, deps);
    deps.setDefinedNameResolvers?.(resolver.resolveDefinedNameValue, resolver.resolveDefinedNameRange, resolver.resolveStructuredRange);
    try {
      for (let pass = 0; pass < 8; pass += 1) {
        let resolvedInPass = 0;
        for (const sheet of workbook.sheets) {
          for (const cell of sheet.cells) {
            if (!cell.formulaText) continue;
            if (cell.resolutionStatus === "unsupported_external") continue;
            if (cell.resolutionStatus === "resolved") continue;
            const reference = deps.parseSimpleFormulaReference(cell.formulaText, sheet.name);
            if (reference) {
              const targetValue = String(resolver.resolveCellValue(reference.sheetName, reference.address) || "").trim();
              if (targetValue) {
                deps.applyResolvedFormulaValue(cell, targetValue, "legacy_resolver");
                resolvedInPass += 1;
                continue;
              }
            }
            let evaluated: string | null = null;
            let evaluatedSource: FormulaResolutionSource = null;
            try {
              const result = deps.tryResolveFormulaExpressionDetailed(
                cell.formulaText,
                sheet.name,
                resolver.resolveCellValue,
                resolver.resolveRangeValues,
                resolver.resolveRangeEntries,
                cell.address
              );
              evaluated = result?.value ?? null;
              evaluatedSource = result?.source ?? null;
            } catch (error) {
              if (!(error instanceof Error) || error.message !== "__FORMULA_UNRESOLVED__") {
                throw error;
              }
            }
            if (evaluated != null) {
              deps.applyResolvedFormulaValue(cell, evaluated, evaluatedSource || "legacy_resolver");
              resolvedInPass += 1;
            }
          }
        }
        if (resolvedInPass === 0) break;
      }
    } finally {
      deps.setDefinedNameResolvers?.(null, null, null);
    }
  }

  const formulaResolverApi = {
    buildFormulaResolver,
    resolveSimpleFormulaReferences
  };

  moduleRegistry.registerModule("formulaResolver", formulaResolverApi);
})();
