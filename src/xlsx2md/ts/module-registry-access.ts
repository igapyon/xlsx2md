(() => {
  type Xlsx2mdModuleRegistry = {
    getModule: <T>(name: string) => T | undefined;
    requireModule: <T>(name: string, errorMessage: string) => T;
    registerModule: <T>(name: string, moduleValue: T) => T;
  };
  type RuntimeEnvApi = {
    ELEMENT_NODE: number;
    TEXT_NODE: number;
    xmlToDocument: (xmlText: string) => Document;
  };
  type MarkdownNormalizeApi = {
    normalizeMarkdownText: (text: string) => string;
    normalizeMarkdownHeadingText: (text: string) => string;
    normalizeMarkdownListItemText: (text: string) => string;
    normalizeMarkdownTableCell: (text: string) => string;
  };
  type MarkdownEscapeApi = {
    escapeMarkdownLineStart: (text: string) => string;
    escapeMarkdownLiteralParts: (text: string) => Array<{ kind: "text" | "escaped"; text: string; rawText: string }>;
    escapeMarkdownLiteralText: (text: string) => string;
  };
  type MarkdownTableEscapeApi = {
    escapeMarkdownTableCell: (text: string) => string;
  };
  type Xlsx2mdRichTextParserModule<TCell> = {
    createRichTextParserApi: (deps: {
      normalizeMarkdownText?: (text: string) => string;
    }) => {
      compactText: (text: string) => string;
      splitTextWithLineBreaks: (text: string) => Array<
        | { kind: "text"; text: string }
        | { kind: "lineBreak" }
        | {
          kind: "styledText";
          parts: Array<{ kind: "text" | "escaped"; text: string; rawText: string }>;
          style: { bold: boolean; italic: boolean; strike: boolean; underline: boolean };
        }
      >;
      createStyledTextToken: (
        text: string,
        style: { bold: boolean; italic: boolean; strike: boolean; underline: boolean }
      ) => {
        kind: "styledText";
        parts: Array<{ kind: "text" | "escaped"; text: string; rawText: string }>;
        style: { bold: boolean; italic: boolean; strike: boolean; underline: boolean };
      };
      tokenizePlainCellText: (text: string) => Array<{ kind: "text"; text: string }>;
      tokenizeGithubCellText: (
        text: string,
        style: { bold: boolean; italic: boolean; strike: boolean; underline: boolean }
      ) => Array<
        | { kind: "lineBreak" }
        | {
          kind: "styledText";
          parts: Array<{ kind: "text" | "escaped"; text: string; rawText: string }>;
          style: { bold: boolean; italic: boolean; strike: boolean; underline: boolean };
        }
      >;
      tokenizeGithubRichTextRuns: (runs: Array<{
        text: string;
        bold: boolean;
        italic: boolean;
        strike: boolean;
        underline: boolean;
      }>) => Array<
        | { kind: "lineBreak" }
        | {
          kind: "styledText";
          parts: Array<{ kind: "text" | "escaped"; text: string; rawText: string }>;
          style: { bold: boolean; italic: boolean; strike: boolean; underline: boolean };
        }
      >;
      tokenizeCellDisplayText: (cell: TCell | undefined, formattingMode?: "plain" | "github") => Array<
        | { kind: "text"; text: string }
        | { kind: "lineBreak" }
        | {
          kind: "styledText";
          parts: Array<{ kind: "text" | "escaped"; text: string; rawText: string }>;
          style: { bold: boolean; italic: boolean; strike: boolean; underline: boolean };
        }
      >;
    };
  };
  type ZipIoApi = {
    unzipEntries: (arrayBuffer: ArrayBuffer) => Promise<Map<string, Uint8Array>>;
    createStoredZip: (entries: Array<{ name: string; data: Uint8Array }>) => Uint8Array;
  };
  type Xlsx2mdDrawingHelperModule = {
    renderShapeSvg?: (shapeNode: Element, anchor: Element, sheetName: string, shapeIndex: number) => {
      filename: string;
      path: string;
      data: Uint8Array;
    } | null;
  };
  type Xlsx2mdNarrativeStructureModule<TNarrativeBlock> = {
    renderNarrativeBlock: (block: TNarrativeBlock) => string;
    isSectionHeadingNarrativeBlock: (block: TNarrativeBlock | null | undefined) => boolean;
  };
  type Xlsx2mdRichTextRendererModule<TCell> = {
    createRichTextRendererApi: (deps: {
      normalizeMarkdownText?: (text: string) => string;
    }) => {
      compactText: (text: string) => string;
      normalizeGithubSegment: (text: string) => string;
      normalizeGithubCellText: (text: string) => string;
      applyTextStyle: (
        text: string,
        style: { bold: boolean; italic: boolean; strike: boolean; underline: boolean }
      ) => string;
      renderStyledTextParts: (parts: Array<{ kind: "text" | "escaped"; text: string; rawText: string }>) => string;
      renderPlainTokens: (tokens: Array<
        | { kind: "text"; text: string }
        | { kind: "lineBreak" }
        | {
          kind: "styledText";
          parts: Array<{ kind: "text" | "escaped"; text: string; rawText: string }>;
          style: { bold: boolean; italic: boolean; strike: boolean; underline: boolean };
        }
      >) => string;
      renderGithubTokens: (tokens: Array<
        | { kind: "text"; text: string }
        | { kind: "lineBreak" }
        | {
          kind: "styledText";
          parts: Array<{ kind: "text" | "escaped"; text: string }>;
          style: { bold: boolean; italic: boolean; strike: boolean; underline: boolean };
        }
      >) => string;
      renderCellDisplayText: (cell: TCell | undefined, formattingMode?: "plain" | "github") => string;
    };
  };
  type Xlsx2mdRichTextPlainFormatterModule = {
    createRichTextPlainFormatterApi: () => {
      renderStyledTextPart: (part: { kind: "text" | "escaped"; text: string; rawText: string }) => string;
      renderStyledTextParts: (parts: Array<{ kind: "text" | "escaped"; text: string; rawText: string }>) => string;
      renderPlainTokens: (tokens: Array<
        | { kind: "text"; text: string }
        | { kind: "lineBreak" }
        | {
          kind: "styledText";
          parts: Array<{ kind: "text" | "escaped"; text: string; rawText: string }>;
          style: { bold: boolean; italic: boolean; strike: boolean; underline: boolean };
        }
      >) => string;
    };
  };
  type Xlsx2mdRichTextGithubFormatterModule = {
    createRichTextGithubFormatterApi: () => {
      applyTextStyle: (
        text: string,
        style: { bold: boolean; italic: boolean; strike: boolean; underline: boolean }
      ) => string;
      renderStyledTextPart: (part: { kind: "text" | "escaped"; text: string; rawText: string }) => string;
      renderStyledTextParts: (parts: Array<{ kind: "text" | "escaped"; text: string; rawText: string }>) => string;
      renderGithubTokens: (tokens: Array<
        | { kind: "text"; text: string }
        | { kind: "lineBreak" }
        | {
          kind: "styledText";
          parts: Array<{ kind: "text" | "escaped"; text: string; rawText: string }>;
          style: { bold: boolean; italic: boolean; strike: boolean; underline: boolean };
        }
      >) => string;
    };
  };
  type Xlsx2mdTableDetectorModule<TSheet, TCell, TMergeRange, TCandidate, TOptions> = {
    detectTableCandidates: (
      sheet: TSheet,
      buildCellMap: (sheet: TSheet) => Map<string, TCell>,
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
    ) => TCandidate[];
    matrixFromCandidate: (
      sheet: TSheet,
      candidate: TCandidate,
      options: TOptions,
      buildCellMap: (sheet: TSheet) => Map<string, TCell>,
      formatCellForMarkdown: (cell: TCell | undefined, options: TOptions) => string
    ) => string[][];
    applyMergeTokens: (
      matrix: string[][],
      merges: TMergeRange[],
      startRow: number,
      startCol: number,
      endRow: number,
      endCol: number
    ) => void;
  };
  type Xlsx2mdMarkdownExportModule<TParsedWorkbook, TMarkdownFile, TExportEntry> = {
    renderMarkdownTable: (rows: string[][], treatFirstRowAsHeader: boolean) => string;
    createOutputFileName: (
      workbookName: string,
      sheetIndex: number,
      sheetName: string,
      outputMode?: "display" | "raw" | "both",
      formattingMode?: "plain" | "github"
    ) => string;
    createSummaryText: (markdownFile: TMarkdownFile) => string;
    createCombinedMarkdownExportFile: (workbook: TParsedWorkbook, markdownFiles: TMarkdownFile[]) => { fileName: string; content: string };
    createExportEntries: (workbook: TParsedWorkbook, markdownFiles: TMarkdownFile[]) => TExportEntry[];
    createWorkbookExportArchive: (workbook: TParsedWorkbook, markdownFiles: TMarkdownFile[]) => Uint8Array;
    normalizeMarkdownLineBreaks: (text: string) => string;
    textEncoder: TextEncoder;
  };
  type Xlsx2mdStylesParserModule<TCellStyleInfo> = {
    BUILTIN_FORMAT_CODES: Record<number, string>;
    hasBorderSide: (side: Element | null) => boolean;
    parseCellStyles: (files: Map<string, Uint8Array>) => TCellStyleInfo[];
  };
  type Xlsx2mdSharedStringsModule = {
    parseSharedStrings: (files: Map<string, Uint8Array>) => Array<{
      text: string;
      runs: Array<{
        text: string;
        bold: boolean;
        italic: boolean;
        strike: boolean;
        underline: boolean;
      }> | null;
    }>;
  };
  type Xlsx2mdWorksheetTablesModule<TParsedTable> = {
    normalizeStructuredTableKey: (value: string) => string;
    parseWorksheetTables: (
      files: Map<string, Uint8Array>,
      worksheetDoc: Document,
      sheetName: string,
      sheetPath: string
    ) => TParsedTable[];
  };
  type Xlsx2mdCellFormatModule<TCell, TCellStyleInfo, TResolutionSource> = {
    formatTextFunctionValue: (value: string, formatText: string) => string | null;
    excelSerialToIsoText: (serial: number) => string;
    formatCellDisplayValue: (rawValue: string, cellStyle: TCellStyleInfo) => string | null;
    applyResolvedFormulaValue: (
      cell: TCell,
      resolvedValue: string,
      resolutionSource?: TResolutionSource
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
  type Xlsx2mdXmlUtilsModule = {
    xmlToDocument: (xmlText: string) => Document;
    getElementsByLocalName: (root: ParentNode, localName: string) => Element[];
    getFirstChildByLocalName: (root: ParentNode, localName: string) => Element | null;
    getDirectChildByLocalName: (root: ParentNode | null, localName: string) => Element | null;
    decodeXmlText: (bytes: Uint8Array) => string;
    getTextContent: (node: Element | null | undefined) => string;
  };
  type Xlsx2mdAddressUtilsModule<TMergeRange> = {
    colToLetters: (col: number) => string;
    lettersToCol: (letters: string) => number;
    parseCellAddress: (address: string) => { row: number; col: number };
    normalizeFormulaAddress: (address: string) => string;
    formatRange: (startRow: number, startCol: number, endRow: number, endCol: number) => string;
    parseRangeRef: (ref: string) => TMergeRange;
    parseRangeAddress: (rawRange: string) => { start: string; end: string } | null;
  };
  type Xlsx2mdRelsParserHelper = {
    normalizeZipPath: (baseFilePath: string, targetPath: string) => string;
    parseRelationships: (files: Map<string, Uint8Array>, relsPath: string, sourcePath: string) => Map<string, string>;
    buildRelsPath: (sourcePath: string) => string;
  };
  type Xlsx2mdRelsParserModule = {
    createRelsParserApi: (deps: Record<string, unknown>) => Xlsx2mdRelsParserHelper;
  };
  type Xlsx2mdFormulaReferenceUtilsHelper = {
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
  type Xlsx2mdFormulaReferenceUtilsModule = {
    createFormulaReferenceUtilsApi: (deps: Record<string, unknown>) => Xlsx2mdFormulaReferenceUtilsHelper;
  };
  type Xlsx2mdSheetMarkdownHelper<TSheet, TCell, TCandidate, TNarrativeBlock, TSectionBlock, TParsedWorkbook, TMarkdownOptions, TMarkdownFile> = {
    buildCellMap: (sheet: TSheet) => Map<string, TCell>;
    formatCellForMarkdown: (cell: TCell | undefined, options: TMarkdownOptions) => string;
    isCellInAnyTable: (row: number, col: number, tables: TCandidate[]) => boolean;
    splitNarrativeRowSegments: (cells: TCell[], options: TMarkdownOptions) => Array<{ startCol: number; values: string[] }>;
    extractNarrativeBlocks: (sheet: TSheet, tables: TCandidate[], options?: TMarkdownOptions) => TNarrativeBlock[];
    extractSectionBlocks: (sheet: TSheet, tables: TCandidate[], narrativeBlocks: TNarrativeBlock[]) => TSectionBlock[];
    convertSheetToMarkdown: (workbook: TParsedWorkbook, sheet: TSheet, options?: TMarkdownOptions) => TMarkdownFile;
    convertWorkbookToMarkdownFiles: (workbook: TParsedWorkbook, options?: TMarkdownOptions) => TMarkdownFile[];
  };
  type Xlsx2mdSheetMarkdownModule<TSheet, TCell, TCandidate, TNarrativeBlock, TSectionBlock, TParsedWorkbook, TMarkdownOptions, TMarkdownFile> = {
    createSheetMarkdownApi: (deps: Record<string, unknown>) => Xlsx2mdSheetMarkdownHelper<TSheet, TCell, TCandidate, TNarrativeBlock, TSectionBlock, TParsedWorkbook, TMarkdownOptions, TMarkdownFile>;
  };
  type Xlsx2mdFormulaEngineHelper<TResolutionSource> = {
    tryResolveFormulaExpressionDetailed: (
      formulaText: string,
      currentSheetName: string,
      resolveCellValue: (sheetName: string, address: string) => string,
      resolveRangeValues?: (sheetName: string, rangeText: string) => number[],
      resolveRangeEntries?: (sheetName: string, rangeText: string) => { rawValues: string[]; numericValues: number[] },
      currentAddress?: string
    ) => { value: string; source: TResolutionSource } | null;
    tryResolveFormulaExpression: (
      formulaText: string,
      currentSheetName: string,
      resolveCellValue: (sheetName: string, address: string) => string,
      resolveRangeValues?: (sheetName: string, rangeText: string) => number[],
      resolveRangeEntries?: (sheetName: string, rangeText: string) => { rawValues: string[]; numericValues: number[] },
      currentAddress?: string
    ) => string | null;
  };
  type Xlsx2mdFormulaEngineModule<TResolutionSource> = {
    createFormulaEngineApi: (deps: Record<string, unknown>) => Xlsx2mdFormulaEngineHelper<TResolutionSource>;
  };
  type Xlsx2mdSheetAssetsModule<TParsedImageAsset, TParsedChartAsset, TParsedShapeAsset, TShapeBlock> = {
    parseDrawingImages: (
      files: Map<string, Uint8Array>,
      sheetName: string,
      sheetPath: string,
      deps: Record<string, unknown>
    ) => TParsedImageAsset[];
    parseDrawingCharts: (
      files: Map<string, Uint8Array>,
      sheetName: string,
      sheetPath: string,
      deps: Record<string, unknown>
    ) => TParsedChartAsset[];
    parseDrawingShapes: (
      files: Map<string, Uint8Array>,
      sheetName: string,
      sheetPath: string,
      deps: Record<string, unknown>
    ) => TParsedShapeAsset[];
    extractShapeBlocks: (
      shapes: TParsedShapeAsset[],
      deps: {
        defaultCellWidthEmu: number;
        defaultCellHeightEmu: number;
        shapeBlockGapXEmu: number;
        shapeBlockGapYEmu: number;
      }
    ) => TShapeBlock[];
    renderHierarchicalRawEntries: (entries: { key: string; value: string }[]) => string[];
  };
  type Xlsx2mdWorksheetParserModule<TCellStyleInfo, TParsedSheet> = {
    parseWorksheet: (
      files: Map<string, Uint8Array>,
      sheetName: string,
      sheetPath: string,
      sheetIndex: number,
      sharedStrings: Array<{
        text: string;
        runs: Array<{
          text: string;
          bold: boolean;
          italic: boolean;
          strike: boolean;
          underline: boolean;
        }> | null;
      }>,
      cellStyles: TCellStyleInfo[],
      deps: Record<string, unknown>
    ) => TParsedSheet;
  };
  type Xlsx2mdWorkbookLoaderModule<TParsedWorkbook> = {
    parseWorkbook: (
      arrayBuffer: ArrayBuffer,
      workbookName: string,
      deps: Record<string, unknown>
    ) => Promise<TParsedWorkbook>;
  };
  type Xlsx2mdFormulaResolverModule<TParsedWorkbook> = {
    resolveSimpleFormulaReferences: (workbook: TParsedWorkbook, deps: Record<string, unknown>) => void;
  };
  type Xlsx2mdFormulaLegacyHelper = {
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
  type Xlsx2mdFormulaLegacyModule = {
    createFormulaLegacyApi: (deps: Record<string, unknown>) => Xlsx2mdFormulaLegacyHelper;
  };
  type Xlsx2mdFormulaAstHelper = {
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
  type Xlsx2mdFormulaAstModule = {
    createFormulaAstApi: (deps: Record<string, unknown>) => Xlsx2mdFormulaAstHelper;
  };

  (globalThis as typeof globalThis & {
    getXlsx2mdModuleRegistry?: () => Xlsx2mdModuleRegistry;
    requireXlsx2mdRuntimeEnv?: () => RuntimeEnvApi;
    requireXlsx2mdMarkdownNormalize?: () => MarkdownNormalizeApi;
    requireXlsx2mdZipIo?: () => ZipIoApi;
    getXlsx2mdDrawingHelperModule?: () => Xlsx2mdDrawingHelperModule | null;
    requireXlsx2mdNarrativeStructureModule?: <TNarrativeBlock>() => Xlsx2mdNarrativeStructureModule<TNarrativeBlock>;
    requireXlsx2mdRichTextPlainFormatterModule?: () => Xlsx2mdRichTextPlainFormatterModule;
    requireXlsx2mdRichTextRendererModule?: <TCell>() => Xlsx2mdRichTextRendererModule<TCell>;
    requireXlsx2mdTableDetectorModule?: <TSheet, TCell, TMergeRange, TCandidate, TOptions>() => Xlsx2mdTableDetectorModule<TSheet, TCell, TMergeRange, TCandidate, TOptions>;
    requireXlsx2mdMarkdownExportModule?: <TParsedWorkbook, TMarkdownFile, TExportEntry>() => Xlsx2mdMarkdownExportModule<TParsedWorkbook, TMarkdownFile, TExportEntry>;
    requireXlsx2mdStylesParserModule?: <TCellStyleInfo>() => Xlsx2mdStylesParserModule<TCellStyleInfo>;
    requireXlsx2mdSharedStringsModule?: () => Xlsx2mdSharedStringsModule;
    requireXlsx2mdWorksheetTablesModule?: <TParsedTable>() => Xlsx2mdWorksheetTablesModule<TParsedTable>;
    requireXlsx2mdCellFormatModule?: <TCell, TCellStyleInfo, TResolutionSource>() => Xlsx2mdCellFormatModule<TCell, TCellStyleInfo, TResolutionSource>;
    requireXlsx2mdXmlUtilsModule?: () => Xlsx2mdXmlUtilsModule;
    requireXlsx2mdAddressUtilsModule?: <TMergeRange>() => Xlsx2mdAddressUtilsModule<TMergeRange>;
    requireXlsx2mdRelsParserModule?: () => Xlsx2mdRelsParserModule;
    requireXlsx2mdFormulaReferenceUtilsModule?: () => Xlsx2mdFormulaReferenceUtilsModule;
    requireXlsx2mdSheetMarkdownModule?: <TSheet, TCell, TCandidate, TNarrativeBlock, TSectionBlock, TParsedWorkbook, TMarkdownOptions, TMarkdownFile>() => Xlsx2mdSheetMarkdownModule<TSheet, TCell, TCandidate, TNarrativeBlock, TSectionBlock, TParsedWorkbook, TMarkdownOptions, TMarkdownFile>;
    requireXlsx2mdFormulaEngineModule?: <TResolutionSource>() => Xlsx2mdFormulaEngineModule<TResolutionSource>;
    requireXlsx2mdSheetAssetsModule?: <TParsedImageAsset, TParsedChartAsset, TParsedShapeAsset, TShapeBlock>() => Xlsx2mdSheetAssetsModule<TParsedImageAsset, TParsedChartAsset, TParsedShapeAsset, TShapeBlock>;
    requireXlsx2mdWorksheetParserModule?: <TCellStyleInfo, TParsedSheet>() => Xlsx2mdWorksheetParserModule<TCellStyleInfo, TParsedSheet>;
    requireXlsx2mdWorkbookLoaderModule?: <TParsedWorkbook>() => Xlsx2mdWorkbookLoaderModule<TParsedWorkbook>;
    requireXlsx2mdFormulaResolverModule?: <TParsedWorkbook>() => Xlsx2mdFormulaResolverModule<TParsedWorkbook>;
    requireXlsx2mdFormulaLegacyModule?: () => Xlsx2mdFormulaLegacyModule;
    requireXlsx2mdFormulaAstModule?: () => Xlsx2mdFormulaAstModule;
  }).getXlsx2mdModuleRegistry = function getXlsx2mdModuleRegistry(): Xlsx2mdModuleRegistry {
    const moduleRegistry = (globalThis as typeof globalThis & {
      __xlsx2mdModuleRegistry?: Xlsx2mdModuleRegistry;
    }).__xlsx2mdModuleRegistry;
    if (!moduleRegistry) {
      throw new Error("xlsx2md module registry is not loaded");
    }
    return moduleRegistry;
  };
  (globalThis as typeof globalThis & {
    requireXlsx2mdRuntimeEnv?: () => RuntimeEnvApi;
  }).requireXlsx2mdRuntimeEnv = function requireXlsx2mdRuntimeEnv(): RuntimeEnvApi {
    return (globalThis as typeof globalThis & {
      getXlsx2mdModuleRegistry: () => Xlsx2mdModuleRegistry;
    }).getXlsx2mdModuleRegistry().requireModule<RuntimeEnvApi>(
      "runtimeEnv",
      "xlsx2md runtime env module is not loaded"
    );
  };
  (globalThis as typeof globalThis & {
    requireXlsx2mdMarkdownNormalize?: () => MarkdownNormalizeApi;
  }).requireXlsx2mdMarkdownNormalize = function requireXlsx2mdMarkdownNormalize(): MarkdownNormalizeApi {
    return (globalThis as typeof globalThis & {
      getXlsx2mdModuleRegistry: () => Xlsx2mdModuleRegistry;
    }).getXlsx2mdModuleRegistry().requireModule<MarkdownNormalizeApi>(
      "markdownNormalize",
      "xlsx2md markdown normalize module is not loaded"
    );
  };
  (globalThis as typeof globalThis & {
    requireXlsx2mdZipIo?: () => ZipIoApi;
  }).requireXlsx2mdZipIo = function requireXlsx2mdZipIo(): ZipIoApi {
    return (globalThis as typeof globalThis & {
      getXlsx2mdModuleRegistry: () => Xlsx2mdModuleRegistry;
    }).getXlsx2mdModuleRegistry().requireModule<ZipIoApi>(
      "zipIo",
      "xlsx2md zip io module is not loaded"
    );
  };
  (globalThis as typeof globalThis & {
    requireXlsx2mdMarkdownEscape?: () => MarkdownEscapeApi;
  }).requireXlsx2mdMarkdownEscape = function requireXlsx2mdMarkdownEscape(): MarkdownEscapeApi {
    return (globalThis as typeof globalThis & {
      getXlsx2mdModuleRegistry: () => Xlsx2mdModuleRegistry;
    }).getXlsx2mdModuleRegistry().requireModule<MarkdownEscapeApi>(
      "markdownEscape",
      "xlsx2md markdown escape module is not loaded"
    );
  };
  (globalThis as typeof globalThis & {
    requireXlsx2mdMarkdownTableEscape?: () => MarkdownTableEscapeApi;
  }).requireXlsx2mdMarkdownTableEscape = function requireXlsx2mdMarkdownTableEscape(): MarkdownTableEscapeApi {
    return (globalThis as typeof globalThis & {
      getXlsx2mdModuleRegistry: () => Xlsx2mdModuleRegistry;
    }).getXlsx2mdModuleRegistry().requireModule<MarkdownTableEscapeApi>(
      "markdownTableEscape",
      "xlsx2md markdown table escape module is not loaded"
    );
  };
  (globalThis as typeof globalThis & {
    requireXlsx2mdRichTextPlainFormatterModule?: () => Xlsx2mdRichTextPlainFormatterModule;
  }).requireXlsx2mdRichTextPlainFormatterModule = function requireXlsx2mdRichTextPlainFormatterModule(): Xlsx2mdRichTextPlainFormatterModule {
    return (globalThis as typeof globalThis & {
      getXlsx2mdModuleRegistry: () => Xlsx2mdModuleRegistry;
    }).getXlsx2mdModuleRegistry().requireModule<Xlsx2mdRichTextPlainFormatterModule>(
      "richTextPlainFormatter",
      "xlsx2md rich text plain formatter module is not loaded"
    );
  };
  (globalThis as typeof globalThis & {
    getXlsx2mdDrawingHelperModule?: () => Xlsx2mdDrawingHelperModule | null;
  }).getXlsx2mdDrawingHelperModule = function getXlsx2mdDrawingHelperModule(): Xlsx2mdDrawingHelperModule | null {
    return (globalThis as typeof globalThis & {
      getXlsx2mdModuleRegistry: () => Xlsx2mdModuleRegistry;
    }).getXlsx2mdModuleRegistry().getModule<Xlsx2mdDrawingHelperModule>("officeDrawing") || null;
  };
  (globalThis as typeof globalThis & {
    requireXlsx2mdNarrativeStructureModule?: <TNarrativeBlock>() => Xlsx2mdNarrativeStructureModule<TNarrativeBlock>;
  }).requireXlsx2mdNarrativeStructureModule = function requireXlsx2mdNarrativeStructureModule<TNarrativeBlock>(): Xlsx2mdNarrativeStructureModule<TNarrativeBlock> {
    return (globalThis as typeof globalThis & {
      getXlsx2mdModuleRegistry: () => Xlsx2mdModuleRegistry;
    }).getXlsx2mdModuleRegistry().requireModule<Xlsx2mdNarrativeStructureModule<TNarrativeBlock>>("narrativeStructure", "xlsx2md narrative structure module is not loaded");
  };
  (globalThis as typeof globalThis & {
    requireXlsx2mdRichTextParserModule?: <TCell>() => Xlsx2mdRichTextParserModule<TCell>;
  }).requireXlsx2mdRichTextParserModule = function requireXlsx2mdRichTextParserModule<TCell>(): Xlsx2mdRichTextParserModule<TCell> {
    return (globalThis as typeof globalThis & {
      getXlsx2mdModuleRegistry: () => Xlsx2mdModuleRegistry;
    }).getXlsx2mdModuleRegistry().requireModule<Xlsx2mdRichTextParserModule<TCell>>(
      "richTextParser",
      "xlsx2md rich text parser module is not loaded"
    );
  };
  (globalThis as typeof globalThis & {
    requireXlsx2mdRichTextGithubFormatterModule?: () => Xlsx2mdRichTextGithubFormatterModule;
  }).requireXlsx2mdRichTextGithubFormatterModule = function requireXlsx2mdRichTextGithubFormatterModule(): Xlsx2mdRichTextGithubFormatterModule {
    return (globalThis as typeof globalThis & {
      getXlsx2mdModuleRegistry: () => Xlsx2mdModuleRegistry;
    }).getXlsx2mdModuleRegistry().requireModule<Xlsx2mdRichTextGithubFormatterModule>(
      "richTextGithubFormatter",
      "xlsx2md rich text github formatter module is not loaded"
    );
  };
  (globalThis as typeof globalThis & {
    requireXlsx2mdRichTextRendererModule?: <TCell>() => Xlsx2mdRichTextRendererModule<TCell>;
  }).requireXlsx2mdRichTextRendererModule = function requireXlsx2mdRichTextRendererModule<TCell>(): Xlsx2mdRichTextRendererModule<TCell> {
    return (globalThis as typeof globalThis & {
      getXlsx2mdModuleRegistry: () => Xlsx2mdModuleRegistry;
    }).getXlsx2mdModuleRegistry().requireModule<Xlsx2mdRichTextRendererModule<TCell>>("richTextRenderer", "xlsx2md rich text renderer module is not loaded");
  };
  (globalThis as typeof globalThis & {
    requireXlsx2mdTableDetectorModule?: <TSheet, TCell, TMergeRange, TCandidate, TOptions>() => Xlsx2mdTableDetectorModule<TSheet, TCell, TMergeRange, TCandidate, TOptions>;
  }).requireXlsx2mdTableDetectorModule = function requireXlsx2mdTableDetectorModule<TSheet, TCell, TMergeRange, TCandidate, TOptions>(): Xlsx2mdTableDetectorModule<TSheet, TCell, TMergeRange, TCandidate, TOptions> {
    return (globalThis as typeof globalThis & {
      getXlsx2mdModuleRegistry: () => Xlsx2mdModuleRegistry;
    }).getXlsx2mdModuleRegistry().requireModule<Xlsx2mdTableDetectorModule<TSheet, TCell, TMergeRange, TCandidate, TOptions>>("tableDetector", "xlsx2md table detector module is not loaded");
  };
  (globalThis as typeof globalThis & {
    requireXlsx2mdMarkdownExportModule?: <TParsedWorkbook, TMarkdownFile, TExportEntry>() => Xlsx2mdMarkdownExportModule<TParsedWorkbook, TMarkdownFile, TExportEntry>;
  }).requireXlsx2mdMarkdownExportModule = function requireXlsx2mdMarkdownExportModule<TParsedWorkbook, TMarkdownFile, TExportEntry>(): Xlsx2mdMarkdownExportModule<TParsedWorkbook, TMarkdownFile, TExportEntry> {
    return (globalThis as typeof globalThis & {
      getXlsx2mdModuleRegistry: () => Xlsx2mdModuleRegistry;
    }).getXlsx2mdModuleRegistry().requireModule<Xlsx2mdMarkdownExportModule<TParsedWorkbook, TMarkdownFile, TExportEntry>>("markdownExport", "xlsx2md markdown export module is not loaded");
  };
  (globalThis as typeof globalThis & {
    requireXlsx2mdStylesParserModule?: <TCellStyleInfo>() => Xlsx2mdStylesParserModule<TCellStyleInfo>;
  }).requireXlsx2mdStylesParserModule = function requireXlsx2mdStylesParserModule<TCellStyleInfo>(): Xlsx2mdStylesParserModule<TCellStyleInfo> {
    return (globalThis as typeof globalThis & {
      getXlsx2mdModuleRegistry: () => Xlsx2mdModuleRegistry;
    }).getXlsx2mdModuleRegistry().requireModule<Xlsx2mdStylesParserModule<TCellStyleInfo>>("stylesParser", "xlsx2md styles parser module is not loaded");
  };
  (globalThis as typeof globalThis & {
    requireXlsx2mdSharedStringsModule?: () => Xlsx2mdSharedStringsModule;
  }).requireXlsx2mdSharedStringsModule = function requireXlsx2mdSharedStringsModule(): Xlsx2mdSharedStringsModule {
    return (globalThis as typeof globalThis & {
      getXlsx2mdModuleRegistry: () => Xlsx2mdModuleRegistry;
    }).getXlsx2mdModuleRegistry().requireModule<Xlsx2mdSharedStringsModule>("sharedStrings", "xlsx2md shared strings module is not loaded");
  };
  (globalThis as typeof globalThis & {
    requireXlsx2mdWorksheetTablesModule?: <TParsedTable>() => Xlsx2mdWorksheetTablesModule<TParsedTable>;
  }).requireXlsx2mdWorksheetTablesModule = function requireXlsx2mdWorksheetTablesModule<TParsedTable>(): Xlsx2mdWorksheetTablesModule<TParsedTable> {
    return (globalThis as typeof globalThis & {
      getXlsx2mdModuleRegistry: () => Xlsx2mdModuleRegistry;
    }).getXlsx2mdModuleRegistry().requireModule<Xlsx2mdWorksheetTablesModule<TParsedTable>>("worksheetTables", "xlsx2md worksheet tables module is not loaded");
  };
  (globalThis as typeof globalThis & {
    requireXlsx2mdCellFormatModule?: <TCell, TCellStyleInfo, TResolutionSource>() => Xlsx2mdCellFormatModule<TCell, TCellStyleInfo, TResolutionSource>;
  }).requireXlsx2mdCellFormatModule = function requireXlsx2mdCellFormatModule<TCell, TCellStyleInfo, TResolutionSource>(): Xlsx2mdCellFormatModule<TCell, TCellStyleInfo, TResolutionSource> {
    return (globalThis as typeof globalThis & {
      getXlsx2mdModuleRegistry: () => Xlsx2mdModuleRegistry;
    }).getXlsx2mdModuleRegistry().requireModule<Xlsx2mdCellFormatModule<TCell, TCellStyleInfo, TResolutionSource>>("cellFormat", "xlsx2md cell format module is not loaded");
  };
  (globalThis as typeof globalThis & {
    requireXlsx2mdXmlUtilsModule?: () => Xlsx2mdXmlUtilsModule;
  }).requireXlsx2mdXmlUtilsModule = function requireXlsx2mdXmlUtilsModule(): Xlsx2mdXmlUtilsModule {
    return (globalThis as typeof globalThis & {
      getXlsx2mdModuleRegistry: () => Xlsx2mdModuleRegistry;
    }).getXlsx2mdModuleRegistry().requireModule<Xlsx2mdXmlUtilsModule>("xmlUtils", "xlsx2md xml utils module is not loaded");
  };
  (globalThis as typeof globalThis & {
    requireXlsx2mdAddressUtilsModule?: <TMergeRange>() => Xlsx2mdAddressUtilsModule<TMergeRange>;
  }).requireXlsx2mdAddressUtilsModule = function requireXlsx2mdAddressUtilsModule<TMergeRange>(): Xlsx2mdAddressUtilsModule<TMergeRange> {
    return (globalThis as typeof globalThis & {
      getXlsx2mdModuleRegistry: () => Xlsx2mdModuleRegistry;
    }).getXlsx2mdModuleRegistry().requireModule<Xlsx2mdAddressUtilsModule<TMergeRange>>("addressUtils", "xlsx2md address utils module is not loaded");
  };
  (globalThis as typeof globalThis & {
    requireXlsx2mdRelsParserModule?: () => Xlsx2mdRelsParserModule;
  }).requireXlsx2mdRelsParserModule = function requireXlsx2mdRelsParserModule(): Xlsx2mdRelsParserModule {
    return (globalThis as typeof globalThis & {
      getXlsx2mdModuleRegistry: () => Xlsx2mdModuleRegistry;
    }).getXlsx2mdModuleRegistry().requireModule<Xlsx2mdRelsParserModule>("relsParser", "xlsx2md rels parser module is not loaded");
  };
  (globalThis as typeof globalThis & {
    requireXlsx2mdFormulaReferenceUtilsModule?: () => Xlsx2mdFormulaReferenceUtilsModule;
  }).requireXlsx2mdFormulaReferenceUtilsModule = function requireXlsx2mdFormulaReferenceUtilsModule(): Xlsx2mdFormulaReferenceUtilsModule {
    return (globalThis as typeof globalThis & {
      getXlsx2mdModuleRegistry: () => Xlsx2mdModuleRegistry;
    }).getXlsx2mdModuleRegistry().requireModule<Xlsx2mdFormulaReferenceUtilsModule>("formulaReferenceUtils", "xlsx2md formula reference utils module is not loaded");
  };
  (globalThis as typeof globalThis & {
    requireXlsx2mdSheetMarkdownModule?: <TSheet, TCell, TCandidate, TNarrativeBlock, TSectionBlock, TParsedWorkbook, TMarkdownOptions, TMarkdownFile>() => Xlsx2mdSheetMarkdownModule<TSheet, TCell, TCandidate, TNarrativeBlock, TSectionBlock, TParsedWorkbook, TMarkdownOptions, TMarkdownFile>;
  }).requireXlsx2mdSheetMarkdownModule = function requireXlsx2mdSheetMarkdownModule<TSheet, TCell, TCandidate, TNarrativeBlock, TSectionBlock, TParsedWorkbook, TMarkdownOptions, TMarkdownFile>(): Xlsx2mdSheetMarkdownModule<TSheet, TCell, TCandidate, TNarrativeBlock, TSectionBlock, TParsedWorkbook, TMarkdownOptions, TMarkdownFile> {
    return (globalThis as typeof globalThis & {
      getXlsx2mdModuleRegistry: () => Xlsx2mdModuleRegistry;
    }).getXlsx2mdModuleRegistry().requireModule<Xlsx2mdSheetMarkdownModule<TSheet, TCell, TCandidate, TNarrativeBlock, TSectionBlock, TParsedWorkbook, TMarkdownOptions, TMarkdownFile>>("sheetMarkdown", "xlsx2md sheet markdown module is not loaded");
  };
  (globalThis as typeof globalThis & {
    requireXlsx2mdFormulaEngineModule?: <TResolutionSource>() => Xlsx2mdFormulaEngineModule<TResolutionSource>;
  }).requireXlsx2mdFormulaEngineModule = function requireXlsx2mdFormulaEngineModule<TResolutionSource>(): Xlsx2mdFormulaEngineModule<TResolutionSource> {
    return (globalThis as typeof globalThis & {
      getXlsx2mdModuleRegistry: () => Xlsx2mdModuleRegistry;
    }).getXlsx2mdModuleRegistry().requireModule<Xlsx2mdFormulaEngineModule<TResolutionSource>>("formulaEngine", "xlsx2md formula engine module is not loaded");
  };
  (globalThis as typeof globalThis & {
    requireXlsx2mdSheetAssetsModule?: <TParsedImageAsset, TParsedChartAsset, TParsedShapeAsset, TShapeBlock>() => Xlsx2mdSheetAssetsModule<TParsedImageAsset, TParsedChartAsset, TParsedShapeAsset, TShapeBlock>;
  }).requireXlsx2mdSheetAssetsModule = function requireXlsx2mdSheetAssetsModule<TParsedImageAsset, TParsedChartAsset, TParsedShapeAsset, TShapeBlock>(): Xlsx2mdSheetAssetsModule<TParsedImageAsset, TParsedChartAsset, TParsedShapeAsset, TShapeBlock> {
    return (globalThis as typeof globalThis & {
      getXlsx2mdModuleRegistry: () => Xlsx2mdModuleRegistry;
    }).getXlsx2mdModuleRegistry().requireModule<Xlsx2mdSheetAssetsModule<TParsedImageAsset, TParsedChartAsset, TParsedShapeAsset, TShapeBlock>>("sheetAssets", "xlsx2md sheet assets module is not loaded");
  };
  (globalThis as typeof globalThis & {
    requireXlsx2mdWorksheetParserModule?: <TCellStyleInfo, TParsedSheet>() => Xlsx2mdWorksheetParserModule<TCellStyleInfo, TParsedSheet>;
  }).requireXlsx2mdWorksheetParserModule = function requireXlsx2mdWorksheetParserModule<TCellStyleInfo, TParsedSheet>(): Xlsx2mdWorksheetParserModule<TCellStyleInfo, TParsedSheet> {
    return (globalThis as typeof globalThis & {
      getXlsx2mdModuleRegistry: () => Xlsx2mdModuleRegistry;
    }).getXlsx2mdModuleRegistry().requireModule<Xlsx2mdWorksheetParserModule<TCellStyleInfo, TParsedSheet>>("worksheetParser", "xlsx2md worksheet parser module is not loaded");
  };
  (globalThis as typeof globalThis & {
    requireXlsx2mdWorkbookLoaderModule?: <TParsedWorkbook>() => Xlsx2mdWorkbookLoaderModule<TParsedWorkbook>;
  }).requireXlsx2mdWorkbookLoaderModule = function requireXlsx2mdWorkbookLoaderModule<TParsedWorkbook>(): Xlsx2mdWorkbookLoaderModule<TParsedWorkbook> {
    return (globalThis as typeof globalThis & {
      getXlsx2mdModuleRegistry: () => Xlsx2mdModuleRegistry;
    }).getXlsx2mdModuleRegistry().requireModule<Xlsx2mdWorkbookLoaderModule<TParsedWorkbook>>("workbookLoader", "xlsx2md workbook loader module is not loaded");
  };
  (globalThis as typeof globalThis & {
    requireXlsx2mdFormulaResolverModule?: <TParsedWorkbook>() => Xlsx2mdFormulaResolverModule<TParsedWorkbook>;
  }).requireXlsx2mdFormulaResolverModule = function requireXlsx2mdFormulaResolverModule<TParsedWorkbook>(): Xlsx2mdFormulaResolverModule<TParsedWorkbook> {
    return (globalThis as typeof globalThis & {
      getXlsx2mdModuleRegistry: () => Xlsx2mdModuleRegistry;
    }).getXlsx2mdModuleRegistry().requireModule<Xlsx2mdFormulaResolverModule<TParsedWorkbook>>("formulaResolver", "xlsx2md formula resolver module is not loaded");
  };
  (globalThis as typeof globalThis & {
    requireXlsx2mdFormulaLegacyModule?: () => Xlsx2mdFormulaLegacyModule;
  }).requireXlsx2mdFormulaLegacyModule = function requireXlsx2mdFormulaLegacyModule(): Xlsx2mdFormulaLegacyModule {
    return (globalThis as typeof globalThis & {
      getXlsx2mdModuleRegistry: () => Xlsx2mdModuleRegistry;
    }).getXlsx2mdModuleRegistry().requireModule<Xlsx2mdFormulaLegacyModule>("formulaLegacy", "xlsx2md formula legacy module is not loaded");
  };
  (globalThis as typeof globalThis & {
    requireXlsx2mdFormulaAstModule?: () => Xlsx2mdFormulaAstModule;
  }).requireXlsx2mdFormulaAstModule = function requireXlsx2mdFormulaAstModule(): Xlsx2mdFormulaAstModule {
    return (globalThis as typeof globalThis & {
      getXlsx2mdModuleRegistry: () => Xlsx2mdModuleRegistry;
    }).getXlsx2mdModuleRegistry().requireModule<Xlsx2mdFormulaAstModule>("formulaAst", "xlsx2md formula ast module is not loaded");
  };
})();
