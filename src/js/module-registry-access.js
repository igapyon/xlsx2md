/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    globalThis.getXlsx2mdModuleRegistry = function getXlsx2mdModuleRegistry() {
        const moduleRegistry = globalThis.__xlsx2mdModuleRegistry;
        if (!moduleRegistry) {
            throw new Error("xlsx2md module registry is not loaded");
        }
        return moduleRegistry;
    };
    globalThis.requireXlsx2mdRuntimeEnv = function requireXlsx2mdRuntimeEnv() {
        return globalThis.getXlsx2mdModuleRegistry().requireModule("runtimeEnv", "xlsx2md runtime env module is not loaded");
    };
    globalThis.requireXlsx2mdMarkdownNormalize = function requireXlsx2mdMarkdownNormalize() {
        return globalThis.getXlsx2mdModuleRegistry().requireModule("markdownNormalize", "xlsx2md markdown normalize module is not loaded");
    };
    globalThis.requireXlsx2mdZipIo = function requireXlsx2mdZipIo() {
        return globalThis.getXlsx2mdModuleRegistry().requireModule("zipIo", "xlsx2md zip io module is not loaded");
    };
    globalThis.requireXlsx2mdTextEncoding = function requireXlsx2mdTextEncoding() {
        return globalThis.getXlsx2mdModuleRegistry().requireModule("textEncoding", "xlsx2md text encoding module is not loaded");
    };
    globalThis.requireXlsx2mdMarkdownEscape = function requireXlsx2mdMarkdownEscape() {
        return globalThis.getXlsx2mdModuleRegistry().requireModule("markdownEscape", "xlsx2md markdown escape module is not loaded");
    };
    globalThis.requireXlsx2mdMarkdownTableEscape = function requireXlsx2mdMarkdownTableEscape() {
        return globalThis.getXlsx2mdModuleRegistry().requireModule("markdownTableEscape", "xlsx2md markdown table escape module is not loaded");
    };
    globalThis.requireXlsx2mdRichTextPlainFormatterModule = function requireXlsx2mdRichTextPlainFormatterModule() {
        return globalThis.getXlsx2mdModuleRegistry().requireModule("richTextPlainFormatter", "xlsx2md rich text plain formatter module is not loaded");
    };
    globalThis.getXlsx2mdDrawingHelperModule = function getXlsx2mdDrawingHelperModule() {
        return globalThis.getXlsx2mdModuleRegistry().getModule("officeDrawing") || null;
    };
    globalThis.requireXlsx2mdNarrativeStructureModule = function requireXlsx2mdNarrativeStructureModule() {
        return globalThis.getXlsx2mdModuleRegistry().requireModule("narrativeStructure", "xlsx2md narrative structure module is not loaded");
    };
    globalThis.requireXlsx2mdRichTextParserModule = function requireXlsx2mdRichTextParserModule() {
        return globalThis.getXlsx2mdModuleRegistry().requireModule("richTextParser", "xlsx2md rich text parser module is not loaded");
    };
    globalThis.requireXlsx2mdRichTextGithubFormatterModule = function requireXlsx2mdRichTextGithubFormatterModule() {
        return globalThis.getXlsx2mdModuleRegistry().requireModule("richTextGithubFormatter", "xlsx2md rich text github formatter module is not loaded");
    };
    globalThis.requireXlsx2mdRichTextRendererModule = function requireXlsx2mdRichTextRendererModule() {
        return globalThis.getXlsx2mdModuleRegistry().requireModule("richTextRenderer", "xlsx2md rich text renderer module is not loaded");
    };
    globalThis.requireXlsx2mdTableDetectorModule = function requireXlsx2mdTableDetectorModule() {
        return globalThis.getXlsx2mdModuleRegistry().requireModule("tableDetector", "xlsx2md table detector module is not loaded");
    };
    globalThis.requireXlsx2mdMarkdownExportModule = function requireXlsx2mdMarkdownExportModule() {
        return globalThis.getXlsx2mdModuleRegistry().requireModule("markdownExport", "xlsx2md markdown export module is not loaded");
    };
    globalThis.requireXlsx2mdStylesParserModule = function requireXlsx2mdStylesParserModule() {
        return globalThis.getXlsx2mdModuleRegistry().requireModule("stylesParser", "xlsx2md styles parser module is not loaded");
    };
    globalThis.requireXlsx2mdSharedStringsModule = function requireXlsx2mdSharedStringsModule() {
        return globalThis.getXlsx2mdModuleRegistry().requireModule("sharedStrings", "xlsx2md shared strings module is not loaded");
    };
    globalThis.requireXlsx2mdWorksheetTablesModule = function requireXlsx2mdWorksheetTablesModule() {
        return globalThis.getXlsx2mdModuleRegistry().requireModule("worksheetTables", "xlsx2md worksheet tables module is not loaded");
    };
    globalThis.requireXlsx2mdCellFormatModule = function requireXlsx2mdCellFormatModule() {
        return globalThis.getXlsx2mdModuleRegistry().requireModule("cellFormat", "xlsx2md cell format module is not loaded");
    };
    globalThis.requireXlsx2mdXmlUtilsModule = function requireXlsx2mdXmlUtilsModule() {
        return globalThis.getXlsx2mdModuleRegistry().requireModule("xmlUtils", "xlsx2md xml utils module is not loaded");
    };
    globalThis.requireXlsx2mdAddressUtilsModule = function requireXlsx2mdAddressUtilsModule() {
        return globalThis.getXlsx2mdModuleRegistry().requireModule("addressUtils", "xlsx2md address utils module is not loaded");
    };
    globalThis.requireXlsx2mdRelsParserModule = function requireXlsx2mdRelsParserModule() {
        return globalThis.getXlsx2mdModuleRegistry().requireModule("relsParser", "xlsx2md rels parser module is not loaded");
    };
    globalThis.requireXlsx2mdFormulaReferenceUtilsModule = function requireXlsx2mdFormulaReferenceUtilsModule() {
        return globalThis.getXlsx2mdModuleRegistry().requireModule("formulaReferenceUtils", "xlsx2md formula reference utils module is not loaded");
    };
    globalThis.requireXlsx2mdSheetMarkdownModule = function requireXlsx2mdSheetMarkdownModule() {
        return globalThis.getXlsx2mdModuleRegistry().requireModule("sheetMarkdown", "xlsx2md sheet markdown module is not loaded");
    };
    globalThis.requireXlsx2mdFormulaEngineModule = function requireXlsx2mdFormulaEngineModule() {
        return globalThis.getXlsx2mdModuleRegistry().requireModule("formulaEngine", "xlsx2md formula engine module is not loaded");
    };
    globalThis.requireXlsx2mdSheetAssetsModule = function requireXlsx2mdSheetAssetsModule() {
        return globalThis.getXlsx2mdModuleRegistry().requireModule("sheetAssets", "xlsx2md sheet assets module is not loaded");
    };
    globalThis.requireXlsx2mdWorksheetParserModule = function requireXlsx2mdWorksheetParserModule() {
        return globalThis.getXlsx2mdModuleRegistry().requireModule("worksheetParser", "xlsx2md worksheet parser module is not loaded");
    };
    globalThis.requireXlsx2mdWorkbookLoaderModule = function requireXlsx2mdWorkbookLoaderModule() {
        return globalThis.getXlsx2mdModuleRegistry().requireModule("workbookLoader", "xlsx2md workbook loader module is not loaded");
    };
    globalThis.requireXlsx2mdFormulaResolverModule = function requireXlsx2mdFormulaResolverModule() {
        return globalThis.getXlsx2mdModuleRegistry().requireModule("formulaResolver", "xlsx2md formula resolver module is not loaded");
    };
    globalThis.requireXlsx2mdFormulaLegacyModule = function requireXlsx2mdFormulaLegacyModule() {
        return globalThis.getXlsx2mdModuleRegistry().requireModule("formulaLegacy", "xlsx2md formula legacy module is not loaded");
    };
    globalThis.requireXlsx2mdFormulaAstModule = function requireXlsx2mdFormulaAstModule() {
        return globalThis.getXlsx2mdModuleRegistry().requireModule("formulaAst", "xlsx2md formula ast module is not loaded");
    };
})();
