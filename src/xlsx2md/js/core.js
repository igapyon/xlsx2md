(() => {
    const EMPTY_BORDERS = {
        top: false,
        bottom: false,
        left: false,
        right: false
    };
    const drawingHelper = globalThis.__xlsx2mdOfficeDrawing || null;
    const markdownNormalizeHelper = globalThis.__xlsx2mdMarkdownNormalize;
    if (!markdownNormalizeHelper) {
        throw new Error("xlsx2md markdown normalize module is not loaded");
    }
    const narrativeStructureHelper = globalThis.__xlsx2mdNarrativeStructure;
    if (!narrativeStructureHelper) {
        throw new Error("xlsx2md narrative structure module is not loaded");
    }
    const tableDetectorHelper = globalThis.__xlsx2mdTableDetector;
    if (!tableDetectorHelper) {
        throw new Error("xlsx2md table detector module is not loaded");
    }
    const markdownExportHelper = globalThis.__xlsx2mdMarkdownExport;
    if (!markdownExportHelper) {
        throw new Error("xlsx2md markdown export module is not loaded");
    }
    const stylesParserHelper = globalThis.__xlsx2mdStylesParser;
    if (!stylesParserHelper) {
        throw new Error("xlsx2md styles parser module is not loaded");
    }
    const sharedStringsHelper = globalThis.__xlsx2mdSharedStrings;
    if (!sharedStringsHelper) {
        throw new Error("xlsx2md shared strings module is not loaded");
    }
    const worksheetTablesHelper = globalThis.__xlsx2mdWorksheetTables;
    if (!worksheetTablesHelper) {
        throw new Error("xlsx2md worksheet tables module is not loaded");
    }
    const cellFormatHelper = globalThis.__xlsx2mdCellFormat;
    if (!cellFormatHelper) {
        throw new Error("xlsx2md cell format module is not loaded");
    }
    const xmlUtilsHelper = globalThis.__xlsx2mdXmlUtils;
    if (!xmlUtilsHelper) {
        throw new Error("xlsx2md xml utils module is not loaded");
    }
    const addressUtilsHelper = globalThis.__xlsx2mdAddressUtils;
    if (!addressUtilsHelper) {
        throw new Error("xlsx2md address utils module is not loaded");
    }
    const relsParserModule = globalThis.__xlsx2mdRelsParser;
    if (!relsParserModule) {
        throw new Error("xlsx2md rels parser module is not loaded");
    }
    const formulaReferenceUtilsModule = globalThis.__xlsx2mdFormulaReferenceUtils;
    if (!formulaReferenceUtilsModule) {
        throw new Error("xlsx2md formula reference utils module is not loaded");
    }
    const sheetMarkdownModule = globalThis.__xlsx2mdSheetMarkdown;
    if (!sheetMarkdownModule) {
        throw new Error("xlsx2md sheet markdown module is not loaded");
    }
    const formulaEngineModule = globalThis.__xlsx2mdFormulaEngine;
    if (!formulaEngineModule) {
        throw new Error("xlsx2md formula engine module is not loaded");
    }
    const sheetAssetsHelper = globalThis.__xlsx2mdSheetAssets;
    if (!sheetAssetsHelper) {
        throw new Error("xlsx2md sheet assets module is not loaded");
    }
    const worksheetParserHelper = globalThis.__xlsx2mdWorksheetParser;
    if (!worksheetParserHelper) {
        throw new Error("xlsx2md worksheet parser module is not loaded");
    }
    const workbookLoaderHelper = globalThis.__xlsx2mdWorkbookLoader;
    if (!workbookLoaderHelper) {
        throw new Error("xlsx2md workbook loader module is not loaded");
    }
    const formulaResolverHelper = globalThis.__xlsx2mdFormulaResolver;
    if (!formulaResolverHelper) {
        throw new Error("xlsx2md formula resolver module is not loaded");
    }
    const formulaLegacyModule = globalThis.__xlsx2mdFormulaLegacy;
    if (!formulaLegacyModule) {
        throw new Error("xlsx2md formula legacy module is not loaded");
    }
    const formulaAstModule = globalThis.__xlsx2mdFormulaAst;
    if (!formulaAstModule) {
        throw new Error("xlsx2md formula ast module is not loaded");
    }
    const zipIoHelper = globalThis.__xlsx2mdZipIo;
    if (!zipIoHelper) {
        throw new Error("xlsx2md zip io module is not loaded");
    }
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
    globalThis.__xlsx2md = {
        parseWorkbook,
        unzipEntries: zipIoHelper.unzipEntries,
        parseRangeRef,
        applyMergeTokens: tableDetectorHelper.applyMergeTokens,
        detectTableCandidates: (sheet) => tableDetectorHelper.detectTableCandidates(sheet, buildCellMap),
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
