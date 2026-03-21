(() => {
  type RichTextStyle = {
    bold: boolean;
    italic: boolean;
    strike: boolean;
    underline: boolean;
  };
  type RichTextRun = RichTextStyle & {
    text: string;
  };
  type SharedStringEntry = {
    text: string;
    runs: RichTextRun[] | null;
  };
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
    textStyle: RichTextStyle;
    richTextRuns: RichTextRun[] | null;
    formulaType: string;
    spillRef: string;
  };

  type CellStyleInfo = {
    borders: BorderFlags;
    numFmtId: number;
    formatCode: string;
    textStyle: RichTextStyle;
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
    sharedStrings: SharedStringEntry[];
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
    formattingMode?: "plain" | "github";
  };

  type MarkdownFile = {
    fileName: string;
    sheetName: string;
    markdown: string;
    summary: {
      outputMode: "display" | "raw" | "both";
      formattingMode: "plain" | "github";
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
  const moduleRegistry = getXlsx2mdModuleRegistry();

  function requireCoreNarrativeStructure() {
    return requireXlsx2mdNarrativeStructureModule<NarrativeBlock>();
  }

  function requireCoreTableDetector() {
    return requireXlsx2mdTableDetectorModule<ParsedSheet, ParsedCell, MergeRange, TableCandidate, MarkdownOptions>();
  }

  function requireCoreMarkdownExport() {
    return requireXlsx2mdMarkdownExportModule<ParsedWorkbook, MarkdownFile, ExportEntry>();
  }

  function requireCoreStylesParser() {
    return requireXlsx2mdStylesParserModule<CellStyleInfo>();
  }

  function requireCoreWorksheetTables() {
    return requireXlsx2mdWorksheetTablesModule<ParsedTable>();
  }

  function requireCoreCellFormat() {
    return requireXlsx2mdCellFormatModule<ParsedCell, CellStyleInfo, FormulaResolutionSource>();
  }

  function requireCoreAddressUtils() {
    return requireXlsx2mdAddressUtilsModule<MergeRange>();
  }

  function requireCoreSheetMarkdown() {
    return requireXlsx2mdSheetMarkdownModule<
      ParsedSheet,
      ParsedCell,
      TableCandidate,
      NarrativeBlock,
      SectionBlock,
      ParsedWorkbook,
      MarkdownOptions,
      MarkdownFile
    >();
  }

  function requireCoreFormulaEngine() {
    return requireXlsx2mdFormulaEngineModule<FormulaResolutionSource>();
  }

  function requireCoreSheetAssets() {
    return requireXlsx2mdSheetAssetsModule<ParsedImageAsset, ParsedChartAsset, ParsedShapeAsset, ShapeBlock>();
  }

  function requireCoreWorksheetParser() {
    return requireXlsx2mdWorksheetParserModule<CellStyleInfo, ParsedSheet>();
  }

  function requireCoreWorkbookLoader() {
    return requireXlsx2mdWorkbookLoaderModule<ParsedWorkbook>();
  }

  function requireCoreFormulaResolver() {
    return requireXlsx2mdFormulaResolverModule<ParsedWorkbook>();
  }

  const drawingHelper = getXlsx2mdDrawingHelperModule();
  const markdownNormalizeHelper = requireXlsx2mdMarkdownNormalize();
  const narrativeStructureHelper = requireCoreNarrativeStructure();
  const tableDetectorHelper = requireCoreTableDetector();
  const markdownExportHelper = requireCoreMarkdownExport();
  const stylesParserHelper = requireCoreStylesParser();
  const sharedStringsHelper = requireXlsx2mdSharedStringsModule();
  const worksheetTablesHelper = requireCoreWorksheetTables();
  const cellFormatHelper = requireCoreCellFormat();
  const xmlUtilsHelper = requireXlsx2mdXmlUtilsModule();
  const addressUtilsHelper = requireCoreAddressUtils();
  const relsParserModule = requireXlsx2mdRelsParserModule();
  const formulaReferenceUtilsModule = requireXlsx2mdFormulaReferenceUtilsModule();
  const sheetMarkdownModule = requireCoreSheetMarkdown();
  const formulaEngineModule = requireCoreFormulaEngine();
  const sheetAssetsHelper = requireCoreSheetAssets();
  const worksheetParserHelper = requireCoreWorksheetParser();
  const workbookLoaderHelper = requireCoreWorkbookLoader();
  const formulaResolverHelper = requireCoreFormulaResolver();
  const formulaLegacyModule = requireXlsx2mdFormulaLegacyModule();
  const formulaAstModule = requireXlsx2mdFormulaAstModule();
  const zipIoHelper = moduleRegistry.requireModule<{
    unzipEntries: (arrayBuffer: ArrayBuffer) => Promise<Map<string, Uint8Array>>;
    createStoredZip: (entries: ExportEntry[]) => Uint8Array;
  }>("zipIo", "xlsx2md zip io module is not loaded");
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
        sharedStrings: SharedStringEntry[],
        cellStyles: CellStyleInfo[]
      ) => worksheetParserHelper.parseWorksheet(files, name, sheetPath, sheetIndex, sharedStrings, cellStyles, worksheetParseDeps),
      postProcessWorkbook: (workbook: ParsedWorkbook) => {
        formulaResolverHelper.resolveSimpleFormulaReferences(workbook, formulaResolverDeps);
      }
    });
  }

  const xlsx2mdApi = {
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

  moduleRegistry.registerModule("xlsx2md", xlsx2mdApi);
})();
