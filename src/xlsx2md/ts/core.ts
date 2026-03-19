(() => {
  type BorderFlags = {
    top: boolean;
    bottom: boolean;
    left: boolean;
    right: boolean;
  };

  type FormulaResolutionStatus = "resolved" | "fallback_formula" | "unsupported_external" | null;
  type FormulaResolutionSource = "cached_value" | "ast_evaluator" | "legacy_resolver" | "formula_text" | "external_unsupported" | null;
  type CachedValueState = "present_nonempty" | "present_empty" | "absent" | null;

  type ParsedCell = {
    address: string;
    row: number;
    col: number;
    valueType: string;
    rawValue: string;
    outputValue: string;
    formulaText: string;
    resolutionStatus: FormulaResolutionStatus;
    resolutionSource: FormulaResolutionSource;
    cachedValueState: CachedValueState;
    styleIndex: number;
    borders: BorderFlags;
    numFmtId: number;
    formatCode: string;
    formulaType: string;
    spillRef: string;
  };

  type CellStyleInfo = {
    borders: BorderFlags;
    numFmtId: number;
    formatCode: string;
  };

  type MergeRange = {
    startRow: number;
    startCol: number;
    endRow: number;
    endCol: number;
    ref: string;
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
    index: number;
    path: string;
    cells: ParsedCell[];
    merges: MergeRange[];
    tables: ParsedTable[];
    images: ParsedImageAsset[];
    charts: ParsedChartAsset[];
    shapes: ParsedShapeAsset[];
    maxRow: number;
    maxCol: number;
  };

  type ParsedImageAsset = {
    sheetName: string;
    filename: string;
    path: string;
    anchor: string;
    data: Uint8Array;
    mediaPath: string;
  };

  type ParsedChartAsset = {
    sheetName: string;
    anchor: string;
    chartPath: string;
    title: string;
    chartType: string;
    series: {
      name: string;
      categoriesRef: string;
      valuesRef: string;
      axis: "primary" | "secondary";
    }[];
  };

  type ParsedShapeAsset = {
    sheetName: string;
    anchor: string;
    name: string;
    kind: string;
    text: string;
    widthEmu: number | null;
    heightEmu: number | null;
    elementName: string;
    anchorElementName: string;
    rawEntries: {
      key: string;
      value: string;
    }[];
    bbox: {
      left: number;
      top: number;
      right: number;
      bottom: number;
    };
    svgFilename: string | null;
    svgPath: string | null;
    svgData: Uint8Array | null;
  };

  type ShapeBlock = {
    startRow: number;
    startCol: number;
    endRow: number;
    endCol: number;
    shapeIndexes: number[];
  };

  type ParsedWorkbook = {
    name: string;
    sheets: ParsedSheet[];
    sharedStrings: string[];
    definedNames: {
      name: string;
      formulaText: string;
      localSheetName: string | null;
    }[];
  };

  type TableCandidate = {
    startRow: number;
    startCol: number;
    endRow: number;
    endCol: number;
    score: number;
    reasonSummary: string[];
  };

  type TableScoreDetail = {
    range: string;
    score: number;
    reasons: string[];
  };

  type FormulaDiagnostic = {
    address: string;
    formulaText: string;
    status: FormulaResolutionStatus;
    source: FormulaResolutionSource;
    outputValue: string;
  };

  type NarrativeBlock = {
    startRow: number;
    startCol: number;
    endRow: number;
    lines: string[];
    items: {
      row: number;
      startCol: number;
      text: string;
      cellValues: string[];
    }[];
  };

  type SectionBlock = {
    startRow: number;
    startCol: number;
    endRow: number;
    endCol: number;
  };

  type MarkdownOptions = {
    treatFirstRowAsHeader?: boolean;
    trimText?: boolean;
    removeEmptyRows?: boolean;
    removeEmptyColumns?: boolean;
    includeShapeDetails?: boolean;
    outputMode?: "display" | "raw" | "both";
  };

  type MarkdownFile = {
    fileName: string;
    sheetName: string;
    markdown: string;
    summary: {
      outputMode: "display" | "raw" | "both";
      sections: number;
      tables: number;
      narrativeBlocks: number;
      merges: number;
      images: number;
      charts: number;
      cells: number;
      tableScores: TableScoreDetail[];
      formulaDiagnostics: FormulaDiagnostic[];
    };
  };

  type ExportEntry = { name: string; data: Uint8Array };

  const EMPTY_BORDERS: BorderFlags = {
    top: false,
    bottom: false,
    left: false,
    right: false
  };

  const drawingHelper = (globalThis as typeof globalThis & {
    __xlsx2mdOfficeDrawing?: {
      renderShapeSvg?: (shapeNode: Element, anchor: Element, sheetName: string, shapeIndex: number) => {
        filename: string;
        path: string;
        data: Uint8Array;
      } | null;
    };
  }).__xlsx2mdOfficeDrawing || null;
  const markdownNormalizeHelper = (globalThis as typeof globalThis & {
    __xlsx2mdMarkdownNormalize?: {
      normalizeMarkdownText: (text: string) => string;
    };
  }).__xlsx2mdMarkdownNormalize;
  if (!markdownNormalizeHelper) {
    throw new Error("xlsx2md markdown normalize module is not loaded");
  }
  const narrativeStructureHelper = (globalThis as typeof globalThis & {
    __xlsx2mdNarrativeStructure?: {
      renderNarrativeBlock: (block: NarrativeBlock) => string;
      isSectionHeadingNarrativeBlock: (block: NarrativeBlock | null | undefined) => boolean;
    };
  }).__xlsx2mdNarrativeStructure;
  if (!narrativeStructureHelper) {
    throw new Error("xlsx2md narrative structure module is not loaded");
  }
  const tableDetectorHelper = (globalThis as typeof globalThis & {
    __xlsx2mdTableDetector?: {
      detectTableCandidates: (
        sheet: ParsedSheet,
        buildCellMap: (sheet: ParsedSheet) => Map<string, ParsedCell>,
        scoreWeights?: {
          minGrid: number;
          borderPresence: number;
          densityHigh: number;
          densityVeryHigh: number;
          headerish: number;
          mergeHeavyPenalty: number;
          prosePenalty: number;
          threshold: number;
        }
      ) => TableCandidate[];
      matrixFromCandidate: (
        sheet: ParsedSheet,
        candidate: TableCandidate,
        options: MarkdownOptions,
        buildCellMap: (sheet: ParsedSheet) => Map<string, ParsedCell>,
        formatCellForMarkdown: (cell: ParsedCell | undefined, options: MarkdownOptions) => string
      ) => string[][];
      applyMergeTokens: (
        matrix: string[][],
        merges: MergeRange[],
        startRow: number,
        startCol: number,
        endRow: number,
        endCol: number
      ) => void;
    };
  }).__xlsx2mdTableDetector;
  if (!tableDetectorHelper) {
    throw new Error("xlsx2md table detector module is not loaded");
  }
  const markdownExportHelper = (globalThis as typeof globalThis & {
    __xlsx2mdMarkdownExport?: {
      renderMarkdownTable: (rows: string[][], treatFirstRowAsHeader: boolean) => string;
      createOutputFileName: (
        workbookName: string,
        sheetIndex: number,
        sheetName: string,
        outputMode?: "display" | "raw" | "both"
      ) => string;
      createSummaryText: (markdownFile: MarkdownFile) => string;
      createCombinedMarkdownExportFile: (workbook: ParsedWorkbook, markdownFiles: MarkdownFile[]) => { fileName: string; content: string };
      createExportEntries: (workbook: ParsedWorkbook, markdownFiles: MarkdownFile[]) => ExportEntry[];
      createWorkbookExportArchive: (workbook: ParsedWorkbook, markdownFiles: MarkdownFile[]) => Uint8Array;
      normalizeMarkdownLineBreaks: (text: string) => string;
      textEncoder: TextEncoder;
    };
  }).__xlsx2mdMarkdownExport;
  if (!markdownExportHelper) {
    throw new Error("xlsx2md markdown export module is not loaded");
  }
  const stylesParserHelper = (globalThis as typeof globalThis & {
    __xlsx2mdStylesParser?: {
      BUILTIN_FORMAT_CODES: Record<number, string>;
      hasBorderSide: (side: Element | null) => boolean;
      parseCellStyles: (files: Map<string, Uint8Array>) => CellStyleInfo[];
    };
  }).__xlsx2mdStylesParser;
  if (!stylesParserHelper) {
    throw new Error("xlsx2md styles parser module is not loaded");
  }
  const sharedStringsHelper = (globalThis as typeof globalThis & {
    __xlsx2mdSharedStrings?: {
      parseSharedStrings: (files: Map<string, Uint8Array>) => string[];
    };
  }).__xlsx2mdSharedStrings;
  if (!sharedStringsHelper) {
    throw new Error("xlsx2md shared strings module is not loaded");
  }
  const worksheetTablesHelper = (globalThis as typeof globalThis & {
    __xlsx2mdWorksheetTables?: {
      normalizeStructuredTableKey: (value: string) => string;
      parseWorksheetTables: (
        files: Map<string, Uint8Array>,
        worksheetDoc: Document,
        sheetName: string,
        sheetPath: string
      ) => ParsedTable[];
    };
  }).__xlsx2mdWorksheetTables;
  if (!worksheetTablesHelper) {
    throw new Error("xlsx2md worksheet tables module is not loaded");
  }
  const cellFormatHelper = (globalThis as typeof globalThis & {
    __xlsx2mdCellFormat?: {
      formatTextFunctionValue: (value: string, formatText: string) => string | null;
      excelSerialToIsoText: (serial: number) => string;
      formatCellDisplayValue: (rawValue: string, cellStyle: CellStyleInfo) => string | null;
      applyResolvedFormulaValue: (
        cell: ParsedCell,
        resolvedValue: string,
        resolutionSource?: FormulaResolutionSource
      ) => void;
      parseDateLikeParts: (value: string) => {
        yyyy: string;
        mm: string;
        dd: string;
        hh: string;
        mi: string;
        ss: string;
      } | null;
      datePartsToExcelSerial: (
        year: number,
        month: number,
        day: number,
        hour?: number,
        minute?: number,
        second?: number
      ) => number | null;
      parseValueFunctionText: (value: string) => number | null;
    };
  }).__xlsx2mdCellFormat;
  if (!cellFormatHelper) {
    throw new Error("xlsx2md cell format module is not loaded");
  }
  const xmlUtilsHelper = (globalThis as typeof globalThis & {
    __xlsx2mdXmlUtils?: {
      xmlToDocument: (xmlText: string) => Document;
      getElementsByLocalName: (root: ParentNode, localName: string) => Element[];
      getFirstChildByLocalName: (root: ParentNode, localName: string) => Element | null;
      getDirectChildByLocalName: (root: ParentNode | null, localName: string) => Element | null;
      decodeXmlText: (bytes: Uint8Array) => string;
      getTextContent: (node: Element | null | undefined) => string;
    };
  }).__xlsx2mdXmlUtils;
  if (!xmlUtilsHelper) {
    throw new Error("xlsx2md xml utils module is not loaded");
  }
  const addressUtilsHelper = (globalThis as typeof globalThis & {
    __xlsx2mdAddressUtils?: {
      colToLetters: (col: number) => string;
      lettersToCol: (letters: string) => number;
      parseCellAddress: (address: string) => { row: number; col: number };
      normalizeFormulaAddress: (address: string) => string;
      formatRange: (startRow: number, startCol: number, endRow: number, endCol: number) => string;
      parseRangeRef: (ref: string) => MergeRange;
      parseRangeAddress: (rawRange: string) => { start: string; end: string } | null;
    };
  }).__xlsx2mdAddressUtils;
  if (!addressUtilsHelper) {
    throw new Error("xlsx2md address utils module is not loaded");
  }
  const relsParserModule = (globalThis as typeof globalThis & {
    __xlsx2mdRelsParser?: {
      createRelsParserApi: (deps: Record<string, unknown>) => {
        normalizeZipPath: (baseFilePath: string, targetPath: string) => string;
        parseRelationships: (files: Map<string, Uint8Array>, relsPath: string, sourcePath: string) => Map<string, string>;
        buildRelsPath: (sourcePath: string) => string;
      };
    };
  }).__xlsx2mdRelsParser;
  if (!relsParserModule) {
    throw new Error("xlsx2md rels parser module is not loaded");
  }
  const formulaReferenceUtilsModule = (globalThis as typeof globalThis & {
    __xlsx2mdFormulaReferenceUtils?: {
      createFormulaReferenceUtilsApi: (deps: Record<string, unknown>) => {
        parseSimpleFormulaReference: (
          formulaText: string,
          currentSheetName: string
        ) => { sheetName: string; address: string } | null;
        parseSheetScopedDefinedNameReference: (
          expression: string,
          currentSheetName: string
        ) => { sheetName: string; name: string } | null;
        normalizeFormulaSheetName: (rawName: string) => string;
        normalizeDefinedNameKey: (name: string) => string;
      };
    };
  }).__xlsx2mdFormulaReferenceUtils;
  if (!formulaReferenceUtilsModule) {
    throw new Error("xlsx2md formula reference utils module is not loaded");
  }
  const sheetMarkdownModule = (globalThis as typeof globalThis & {
    __xlsx2mdSheetMarkdown?: {
      createSheetMarkdownApi: (deps: Record<string, unknown>) => {
        buildCellMap: (sheet: ParsedSheet) => Map<string, ParsedCell>;
        formatCellForMarkdown: (cell: ParsedCell | undefined, options: MarkdownOptions) => string;
        isCellInAnyTable: (row: number, col: number, tables: TableCandidate[]) => boolean;
        splitNarrativeRowSegments: (cells: ParsedCell[], options: MarkdownOptions) => Array<{ startCol: number; values: string[] }>;
        extractNarrativeBlocks: (sheet: ParsedSheet, tables: TableCandidate[], options?: MarkdownOptions) => NarrativeBlock[];
        extractSectionBlocks: (sheet: ParsedSheet, tables: TableCandidate[], narrativeBlocks: NarrativeBlock[]) => SectionBlock[];
        convertSheetToMarkdown: (workbook: ParsedWorkbook, sheet: ParsedSheet, options?: MarkdownOptions) => MarkdownFile;
        convertWorkbookToMarkdownFiles: (workbook: ParsedWorkbook, options?: MarkdownOptions) => MarkdownFile[];
      };
    };
  }).__xlsx2mdSheetMarkdown;
  if (!sheetMarkdownModule) {
    throw new Error("xlsx2md sheet markdown module is not loaded");
  }
  const formulaEngineModule = (globalThis as typeof globalThis & {
    __xlsx2mdFormulaEngine?: {
      createFormulaEngineApi: (deps: Record<string, unknown>) => {
        tryResolveFormulaExpressionDetailed: (
          formulaText: string,
          currentSheetName: string,
          resolveCellValue: (sheetName: string, address: string) => string,
          resolveRangeValues?: (sheetName: string, rangeText: string) => number[],
          resolveRangeEntries?: (sheetName: string, rangeText: string) => { rawValues: string[]; numericValues: number[] },
          currentAddress?: string
        ) => { value: string; source: FormulaResolutionSource } | null;
        tryResolveFormulaExpression: (
          formulaText: string,
          currentSheetName: string,
          resolveCellValue: (sheetName: string, address: string) => string,
          resolveRangeValues?: (sheetName: string, rangeText: string) => number[],
          resolveRangeEntries?: (sheetName: string, rangeText: string) => { rawValues: string[]; numericValues: number[] },
          currentAddress?: string
        ) => string | null;
      };
    };
  }).__xlsx2mdFormulaEngine;
  if (!formulaEngineModule) {
    throw new Error("xlsx2md formula engine module is not loaded");
  }
  const sheetAssetsHelper = (globalThis as typeof globalThis & {
    __xlsx2mdSheetAssets?: {
      parseDrawingImages: (
        files: Map<string, Uint8Array>,
        sheetName: string,
        sheetPath: string,
        deps: Record<string, unknown>
      ) => ParsedImageAsset[];
      parseDrawingCharts: (
        files: Map<string, Uint8Array>,
        sheetName: string,
        sheetPath: string,
        deps: Record<string, unknown>
      ) => ParsedChartAsset[];
      parseDrawingShapes: (
        files: Map<string, Uint8Array>,
        sheetName: string,
        sheetPath: string,
        deps: Record<string, unknown>
      ) => ParsedShapeAsset[];
      extractShapeBlocks: (
        shapes: ParsedShapeAsset[],
        deps: {
          defaultCellWidthEmu: number;
          defaultCellHeightEmu: number;
          shapeBlockGapXEmu: number;
          shapeBlockGapYEmu: number;
        }
      ) => ShapeBlock[];
      renderHierarchicalRawEntries: (entries: { key: string; value: string }[]) => string[];
    };
  }).__xlsx2mdSheetAssets;
  if (!sheetAssetsHelper) {
    throw new Error("xlsx2md sheet assets module is not loaded");
  }
  const worksheetParserHelper = (globalThis as typeof globalThis & {
    __xlsx2mdWorksheetParser?: {
      parseWorksheet: (
        files: Map<string, Uint8Array>,
        sheetName: string,
        sheetPath: string,
        sheetIndex: number,
        sharedStrings: string[],
        cellStyles: CellStyleInfo[],
        deps: Record<string, unknown>
      ) => ParsedSheet;
    };
  }).__xlsx2mdWorksheetParser;
  if (!worksheetParserHelper) {
    throw new Error("xlsx2md worksheet parser module is not loaded");
  }
  const workbookLoaderHelper = (globalThis as typeof globalThis & {
    __xlsx2mdWorkbookLoader?: {
      parseWorkbook: (
        arrayBuffer: ArrayBuffer,
        workbookName: string,
        deps: Record<string, unknown>
      ) => Promise<ParsedWorkbook>;
    };
  }).__xlsx2mdWorkbookLoader;
  if (!workbookLoaderHelper) {
    throw new Error("xlsx2md workbook loader module is not loaded");
  }
  const formulaResolverHelper = (globalThis as typeof globalThis & {
    __xlsx2mdFormulaResolver?: {
      resolveSimpleFormulaReferences: (workbook: ParsedWorkbook, deps: Record<string, unknown>) => void;
    };
  }).__xlsx2mdFormulaResolver;
  if (!formulaResolverHelper) {
    throw new Error("xlsx2md formula resolver module is not loaded");
  }
  const formulaLegacyModule = (globalThis as typeof globalThis & {
    __xlsx2mdFormulaLegacy?: {
      createFormulaLegacyApi: (deps: Record<string, unknown>) => {
        tryResolveFormulaExpressionLegacy: (
          normalized: string,
          currentSheetName: string,
          resolveCellValue: (sheetName: string, address: string) => string,
          resolveRangeValues?: (sheetName: string, rangeText: string) => number[],
          resolveRangeEntries?: (sheetName: string, rangeText: string) => { rawValues: string[]; numericValues: number[] }
        ) => string | null;
        findTopLevelOperatorIndex: (expression: string, operator: string) => number;
        parseWholeFunctionCall: (expression: string, allowedNames: string[]) => { name: string; argsText: string } | null;
        splitFormulaArguments: (argText: string) => string[];
        parseQualifiedRangeReference: (argText: string, currentSheetName: string) => { sheetName: string; start: string; end: string } | null;
        resolveScalarFormulaValue: (
          expression: string,
          currentSheetName: string,
          resolveCellValue: (sheetName: string, address: string) => string,
          resolveRangeValues?: (sheetName: string, rangeText: string) => number[],
          resolveRangeEntries?: (sheetName: string, rangeText: string) => { rawValues: string[]; numericValues: number[] }
        ) => string | null;
      };
    };
  }).__xlsx2mdFormulaLegacy;
  if (!formulaLegacyModule) {
    throw new Error("xlsx2md formula legacy module is not loaded");
  }
  const formulaAstModule = (globalThis as typeof globalThis & {
    __xlsx2mdFormulaAst?: {
      createFormulaAstApi: (deps: Record<string, unknown>) => {
        tryResolveFormulaExpressionWithAst: (
          expression: string,
          currentSheetName: string,
          resolveCellValue: (sheetName: string, address: string) => string,
          resolveDefinedNameScalarValue: ((sheetName: string, name: string) => string | null) | null,
          resolveDefinedNameRangeRef: ((sheetName: string, name: string) => { sheetName: string; start: string; end: string } | null) | null,
          resolveStructuredRangeRef: ((sheetName: string, text: string) => { sheetName: string; start: string; end: string } | null) | null,
          resolveSpillRange: ((sheetName: string, ref: string) => { sheetName: string; start: string; end: string } | null),
          resolveRangeEntries?: (sheetName: string, rangeText: string) => { rawValues: string[]; numericValues: number[] },
          currentAddress?: string
        ) => string | null;
      };
    };
  }).__xlsx2mdFormulaAst;
  if (!formulaAstModule) {
    throw new Error("xlsx2md formula ast module is not loaded");
  }
  const zipIoHelper = (globalThis as typeof globalThis & {
    __xlsx2mdZipIo?: {
      unzipEntries: (arrayBuffer: ArrayBuffer) => Promise<Map<string, Uint8Array>>;
      createStoredZip: (entries: ExportEntry[]) => Uint8Array;
    };
  }).__xlsx2mdZipIo;
  if (!zipIoHelper) {
    throw new Error("xlsx2md zip io module is not loaded");
  }
  let resolveDefinedNameScalarValue: ((sheetName: string, name: string) => string | null) | null = null;
  let resolveDefinedNameRangeRef: ((sheetName: string, name: string) => { sheetName: string; start: string; end: string } | null) | null = null;
  let resolveStructuredRangeRef: ((sheetName: string, text: string) => { sheetName: string; start: string; end: string } | null) | null = null;
  const DEFAULT_CELL_WIDTH_EMU = 609600;
  const DEFAULT_CELL_HEIGHT_EMU = 190500;
  const SHAPE_BLOCK_GAP_X_EMU = DEFAULT_CELL_WIDTH_EMU * 4;
  const SHAPE_BLOCK_GAP_Y_EMU = DEFAULT_CELL_HEIGHT_EMU * 6;
  const {
    colToLetters,
    lettersToCol,
    parseCellAddress,
    normalizeFormulaAddress,
    formatRange,
    parseRangeRef,
    parseRangeAddress
  } = addressUtilsHelper;
  const {
    xmlToDocument,
    getElementsByLocalName,
    getFirstChildByLocalName,
    getDirectChildByLocalName,
    decodeXmlText,
    getTextContent
  } = xmlUtilsHelper;
  const relsParserHelper = relsParserModule.createRelsParserApi({
    xmlToDocument,
    decodeXmlText
  });
  const {
    normalizeZipPath,
    parseRelationships,
    buildRelsPath
  } = relsParserHelper;
  const formulaReferenceUtilsHelper = formulaReferenceUtilsModule.createFormulaReferenceUtilsApi({
    normalizeFormulaAddress
  });
  const {
    parseSimpleFormulaReference,
    parseSheetScopedDefinedNameReference,
    normalizeFormulaSheetName,
    normalizeDefinedNameKey
  } = formulaReferenceUtilsHelper;
  const formulaLegacyHelper = formulaLegacyModule.createFormulaLegacyApi({
    normalizeFormulaSheetName,
    normalizeFormulaAddress,
    parseSimpleFormulaReference,
    parseSheetScopedDefinedNameReference,
    parseRangeAddress,
    parseCellAddress,
    colToLetters,
    tryResolveFormulaExpression: (
      formulaText: string,
      currentSheetName: string,
      resolveCellValue: (sheetName: string, address: string) => string,
      resolveRangeValues?: (sheetName: string, rangeText: string) => number[],
      resolveRangeEntries?: (sheetName: string, rangeText: string) => { rawValues: string[]; numericValues: number[] },
      currentAddress?: string
    ) => tryResolveFormulaExpression(
      formulaText,
      currentSheetName,
      resolveCellValue,
      resolveRangeValues,
      resolveRangeEntries,
      currentAddress
    ),
    getDefinedNameScalarValue: () => resolveDefinedNameScalarValue,
    getDefinedNameRangeRef: () => resolveDefinedNameRangeRef,
    getStructuredRangeRef: () => resolveStructuredRangeRef,
    cellFormat: cellFormatHelper
  });
  const formulaAstHelper = formulaAstModule.createFormulaAstApi({
    normalizeFormulaAddress,
    parseSheetScopedDefinedNameReference,
    parseRangeAddress,
    parseCellAddress
  });
  const sheetMarkdownHelper = sheetMarkdownModule.createSheetMarkdownApi({
    renderNarrativeBlock: narrativeStructureHelper.renderNarrativeBlock,
    detectTableCandidates: tableDetectorHelper.detectTableCandidates,
    matrixFromCandidate: tableDetectorHelper.matrixFromCandidate,
    renderMarkdownTable: markdownExportHelper.renderMarkdownTable,
    createOutputFileName: markdownExportHelper.createOutputFileName,
    extractShapeBlocks: sheetAssetsHelper.extractShapeBlocks,
    renderHierarchicalRawEntries: sheetAssetsHelper.renderHierarchicalRawEntries,
    parseCellAddress,
    formatRange,
    colToLetters,
    normalizeMarkdownText: markdownNormalizeHelper.normalizeMarkdownText,
    defaultCellWidthEmu: DEFAULT_CELL_WIDTH_EMU,
    defaultCellHeightEmu: DEFAULT_CELL_HEIGHT_EMU,
    shapeBlockGapXEmu: SHAPE_BLOCK_GAP_X_EMU,
    shapeBlockGapYEmu: SHAPE_BLOCK_GAP_Y_EMU
  });
  const formulaEngineHelper = formulaEngineModule.createFormulaEngineApi({
    getDefinedNameScalarValue: () => resolveDefinedNameScalarValue,
    tryResolveFormulaExpressionWithAst: (
      expression: string,
      currentSheetName: string,
      resolveCellValue: (sheetName: string, address: string) => string,
      resolveRangeEntries?: (sheetName: string, rangeText: string) => { rawValues: string[]; numericValues: number[] },
      currentAddress?: string
    ) => tryResolveFormulaExpressionWithAst(
      expression,
      currentSheetName,
      resolveCellValue,
      resolveRangeEntries,
      currentAddress
    ),
    tryResolveFormulaExpressionLegacy: (
      normalized: string,
      currentSheetName: string,
      resolveCellValue: (sheetName: string, address: string) => string,
      resolveRangeValues?: (sheetName: string, rangeText: string) => number[],
      resolveRangeEntries?: (sheetName: string, rangeText: string) => { rawValues: string[]; numericValues: number[] }
    ) => tryResolveFormulaExpressionLegacy(
      normalized,
      currentSheetName,
      resolveCellValue,
      resolveRangeValues,
      resolveRangeEntries
    )
  });
  const {
    findTopLevelOperatorIndex,
    parseWholeFunctionCall,
    splitFormulaArguments,
    parseQualifiedRangeReference,
    resolveScalarFormulaValue
  } = formulaLegacyHelper;
  const {
    buildCellMap,
    formatCellForMarkdown,
    isCellInAnyTable,
    extractNarrativeBlocks,
    splitNarrativeRowSegments,
    extractSectionBlocks,
    convertSheetToMarkdown,
    convertWorkbookToMarkdownFiles
  } = sheetMarkdownHelper;
  const tryResolveFormulaExpressionLegacy = formulaLegacyHelper.tryResolveFormulaExpressionLegacy;
  const tryResolveFormulaExpressionDetailed = formulaEngineHelper.tryResolveFormulaExpressionDetailed;
  const tryResolveFormulaExpression = formulaEngineHelper.tryResolveFormulaExpression;
  const tryResolveFormulaExpressionWithAst = (
    expression: string,
    currentSheetName: string,
    resolveCellValue: (sheetName: string, address: string) => string,
    resolveRangeEntries?: (sheetName: string, rangeText: string) => { rawValues: string[]; numericValues: number[] },
    currentAddress?: string
  ): string | null => formulaAstHelper.tryResolveFormulaExpressionWithAst(
    expression,
    currentSheetName,
    resolveCellValue,
    resolveDefinedNameScalarValue,
    resolveDefinedNameRangeRef,
    resolveStructuredRangeRef,
    resolveSpillRange,
    resolveRangeEntries,
    currentAddress
  );

  function resolveSpillRange(_sheetName: string, _ref: string): { sheetName: string; start: string; end: string } | null {
    return null;
  }

  function createWorksheetParseDeps() {
    return {
      EMPTY_BORDERS,
      xmlToDocument,
      decodeXmlText,
      getTextContent,
      parseCellAddress,
      parseRangeRef,
      parseWorksheetTables: worksheetTablesHelper.parseWorksheetTables,
      parseDrawingImages: sheetAssetsHelper.parseDrawingImages,
      parseDrawingCharts: sheetAssetsHelper.parseDrawingCharts,
      parseDrawingShapes: sheetAssetsHelper.parseDrawingShapes,
      formatCellDisplayValue: cellFormatHelper.formatCellDisplayValue,
      buildAssetDeps: () => ({
        parseRelationships,
        buildRelsPath,
        xmlToDocument,
        decodeXmlText,
        getElementsByLocalName,
        getFirstChildByLocalName,
        getDirectChildByLocalName,
        getTextContent,
        colToLetters,
        drawingHelper,
        defaultCellWidthEmu: DEFAULT_CELL_WIDTH_EMU,
        defaultCellHeightEmu: DEFAULT_CELL_HEIGHT_EMU,
        shapeBlockGapXEmu: SHAPE_BLOCK_GAP_X_EMU,
        shapeBlockGapYEmu: SHAPE_BLOCK_GAP_Y_EMU
      }),
      lettersToCol,
      colToLetters
    };
  }

  function createFormulaResolverDeps() {
    return {
      normalizeStructuredTableKey: worksheetTablesHelper.normalizeStructuredTableKey,
      normalizeFormulaSheetName,
      normalizeDefinedNameKey,
      normalizeFormulaAddress,
      parseSimpleFormulaReference,
      resolveScalarFormulaValue,
      parseQualifiedRangeReference,
      findTopLevelOperatorIndex,
      parseWholeFunctionCall,
      splitFormulaArguments,
      parseCellAddress,
      colToLetters,
      parseRangeAddress,
      tryResolveFormulaExpressionDetailed,
      applyResolvedFormulaValue: cellFormatHelper.applyResolvedFormulaValue,
      setDefinedNameResolvers: (
        scalar: ((sheetName: string, name: string) => string | null) | null,
        range: ((sheetName: string, name: string) => { sheetName: string; start: string; end: string } | null) | null,
        structured: ((sheetName: string, text: string) => { sheetName: string; start: string; end: string } | null) | null
      ) => {
        resolveDefinedNameScalarValue = scalar;
        resolveDefinedNameRangeRef = range;
        resolveStructuredRangeRef = structured;
      }
    };
  }

  async function parseWorkbook(arrayBuffer: ArrayBuffer, workbookName = "workbook.xlsx"): Promise<ParsedWorkbook> {
    const worksheetParseDeps = createWorksheetParseDeps();
    const formulaResolverDeps = createFormulaResolverDeps();
    return workbookLoaderHelper.parseWorkbook(arrayBuffer, workbookName, {
      unzipEntries: zipIoHelper.unzipEntries,
      parseSharedStrings: sharedStringsHelper.parseSharedStrings,
      parseCellStyles: stylesParserHelper.parseCellStyles,
      parseRelationships,
      xmlToDocument,
      decodeXmlText,
      getTextContent,
      parseWorksheet: (
        files: Map<string, Uint8Array>,
        name: string,
        sheetPath: string,
        sheetIndex: number,
        sharedStrings: string[],
        cellStyles: CellStyleInfo[]
      ) => worksheetParserHelper.parseWorksheet(files, name, sheetPath, sheetIndex, sharedStrings, cellStyles, worksheetParseDeps),
      postProcessWorkbook: (workbook: ParsedWorkbook) => {
        formulaResolverHelper.resolveSimpleFormulaReferences(workbook, formulaResolverDeps);
      }
    });
  }

  (globalThis as typeof globalThis & {
    __xlsx2md?: Record<string, unknown>;
  }).__xlsx2md = {
    parseWorkbook,
    unzipEntries: zipIoHelper.unzipEntries,
    parseRangeRef,
    applyMergeTokens: tableDetectorHelper.applyMergeTokens,
    detectTableCandidates: (sheet: ParsedSheet) => tableDetectorHelper.detectTableCandidates(sheet, buildCellMap),
    extractNarrativeBlocks,
    convertSheetToMarkdown,
    convertWorkbookToMarkdownFiles,
    createSummaryText: markdownExportHelper.createSummaryText,
    createCombinedMarkdownExportFile: markdownExportHelper.createCombinedMarkdownExportFile,
    createExportEntries: markdownExportHelper.createExportEntries,
    createWorkbookExportArchive: markdownExportHelper.createWorkbookExportArchive,
    formatRange,
    colToLetters,
    lettersToCol,
    textEncoder: markdownExportHelper.textEncoder
  };
})();
