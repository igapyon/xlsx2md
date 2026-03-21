(() => {
    const moduleRegistry = getXlsx2mdModuleRegistry();
    const markdownNormalizeHelper = requireXlsx2mdMarkdownNormalize();
    function escapeMarkdownTableCell(text) {
        return markdownNormalizeHelper.escapeMarkdownPipes(markdownNormalizeHelper.normalizeMarkdownText(text));
    }
    const markdownTableEscapeApi = {
        escapeMarkdownTableCell
    };
    moduleRegistry.registerModule("markdownTableEscape", markdownTableEscapeApi);
})();
