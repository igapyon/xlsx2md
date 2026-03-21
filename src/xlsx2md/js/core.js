(() => {
    const EMPTY_BORDERS = {
        top: false,
        bottom: false,
        left: false,
        right: false
    };
    const moduleRegistry = getXlsx2mdModuleRegistry();
    function requireCoreNarrativeStructure() {
        return requireXlsx2mdNarrativeStructureModule();
    }
    function requireCoreTableDetector() {
        return requireXlsx2mdTableDetectorModule();
    }
    function requireCoreMarkdownExport() {
        return requireXlsx2mdMarkdownExportModule();
    }
    function requireCoreStylesParser() {
        return requireXlsx2mdStylesParserModule();
    }
    function requireCoreWorksheetTables() {
        return requireXlsx2mdWorksheetTablesModule();
    }
    function requireCoreCellFormat() {
        return requireXlsx2mdCellFormatModule();
    }
    function requireCoreAddressUtils() {
        return requireXlsx2mdAddressUtilsModule();
    }
    function requireCoreSheetMarkdown() {
        return requireXlsx2mdSheetMarkdownModule();
    }
    function requireCoreFormulaEngine() {
        return requireXlsx2mdFormulaEngineModule();
    }
    function requireCoreSheetAssets() {
        return requireXlsx2mdSheetAssetsModule();
    }
    function requireCoreWorksheetParser() {
        return requireXlsx2mdWorksheetParserModule();
    }
    function requireCoreWorkbookLoader() {
        return requireXlsx2mdWorkbookLoaderModule();
    }
    function requireCoreFormulaResolver() {
        return requireXlsx2mdFormulaResolverModule();
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
    const zipIoHelper = moduleRegistry.requireModule("zipIo", "xlsx2md zip io module is not loaded");
    let resolveDefinedNameScalarValue = null;
    let resolveDefinedNameRangeRef = null;
    let resolveStructuredRangeRef = null;
    const DEFAULT_CELL_WIDTH_EMU = 609600;
    const DEFAULT_CELL_HEIGHT_EMU = 190500;
    const SHAPE_BLOCK_GAP_X_EMU = DEFAULT_CELL_WIDTH_EMU * 4;
    const SHAPE_BLOCK_GAP_Y_EMU = DEFAULT_CELL_HEIGHT_EMU * 6;
    const { colToLetters, lettersToCol, parseCellAddress, normalizeFormulaAddress, formatRange, parseRangeRef, parseRangeAddress } = addressUtilsHelper;
    const { xmlToDocument, getElementsByLocalName, getFirstChildByLocalName, getDirectChildByLocalName, decodeXmlText, getTextContent } = xmlUtilsHelper;
    const relsParserHelper = relsParserModule.createRelsParserApi({
        xmlToDocument,
        decodeXmlText
    });
    const { normalizeZipPath, parseRelationships, buildRelsPath } = relsParserHelper;
    const formulaReferenceUtilsHelper = formulaReferenceUtilsModule.createFormulaReferenceUtilsApi({
        normalizeFormulaAddress
    });
    const { parseSimpleFormulaReference, parseSheetScopedDefinedNameReference, normalizeFormulaSheetName, normalizeDefinedNameKey } = formulaReferenceUtilsHelper;
    const formulaLegacyHelper = formulaLegacyModule.createFormulaLegacyApi({
        normalizeFormulaSheetName,
        normalizeFormulaAddress,
        parseSimpleFormulaReference,
        parseSheetScopedDefinedNameReference,
        parseRangeAddress,
        parseCellAddress,
        colToLetters,
        tryResolveFormulaExpression: (formulaText, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries, currentAddress) => tryResolveFormulaExpression(formulaText, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries, currentAddress),
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
        detectTableCandidates: (sheet, buildCellMapForSheet, tableDetectionMode = "balanced") => tableDetectorHelper.detectTableCandidates(sheet, buildCellMapForSheet, undefined, tableDetectionMode),
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
        tryResolveFormulaExpressionWithAst: (expression, currentSheetName, resolveCellValue, resolveRangeEntries, currentAddress) => tryResolveFormulaExpressionWithAst(expression, currentSheetName, resolveCellValue, resolveRangeEntries, currentAddress),
        tryResolveFormulaExpressionLegacy: (normalized, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries) => tryResolveFormulaExpressionLegacy(normalized, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries)
    });
    const { findTopLevelOperatorIndex, parseWholeFunctionCall, splitFormulaArguments, parseQualifiedRangeReference, resolveScalarFormulaValue } = formulaLegacyHelper;
    const { buildCellMap, formatCellForMarkdown, isCellInAnyTable, extractNarrativeBlocks, splitNarrativeRowSegments, extractSectionBlocks, convertSheetToMarkdown, convertWorkbookToMarkdownFiles } = sheetMarkdownHelper;
    const tryResolveFormulaExpressionLegacy = formulaLegacyHelper.tryResolveFormulaExpressionLegacy;
    const tryResolveFormulaExpressionDetailed = formulaEngineHelper.tryResolveFormulaExpressionDetailed;
    const tryResolveFormulaExpression = formulaEngineHelper.tryResolveFormulaExpression;
    const tryResolveFormulaExpressionWithAst = (expression, currentSheetName, resolveCellValue, resolveRangeEntries, currentAddress) => formulaAstHelper.tryResolveFormulaExpressionWithAst(expression, currentSheetName, resolveCellValue, resolveDefinedNameScalarValue, resolveDefinedNameRangeRef, resolveStructuredRangeRef, resolveSpillRange, resolveRangeEntries, currentAddress);
    function resolveSpillRange(_sheetName, _ref) {
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
            setDefinedNameResolvers: (scalar, range, structured) => {
                resolveDefinedNameScalarValue = scalar;
                resolveDefinedNameRangeRef = range;
                resolveStructuredRangeRef = structured;
            }
        };
    }
    async function parseWorkbook(arrayBuffer, workbookName = "workbook.xlsx") {
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
            parseWorksheet: (files, name, sheetPath, sheetIndex, sharedStrings, cellStyles) => worksheetParserHelper.parseWorksheet(files, name, sheetPath, sheetIndex, sharedStrings, cellStyles, worksheetParseDeps),
            postProcessWorkbook: (workbook) => {
                formulaResolverHelper.resolveSimpleFormulaReferences(workbook, formulaResolverDeps);
            }
        });
    }
    const xlsx2mdApi = {
        parseWorkbook,
        unzipEntries: zipIoHelper.unzipEntries,
        parseRangeRef,
        applyMergeTokens: tableDetectorHelper.applyMergeTokens,
        detectTableCandidates: (sheet, tableDetectionMode = "balanced") => tableDetectorHelper.detectTableCandidates(sheet, buildCellMap, undefined, tableDetectionMode),
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
