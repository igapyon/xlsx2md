(() => {
  const moduleRegistry = getXlsx2mdModuleRegistry();
  const markdownNormalizeHelper = requireXlsx2mdMarkdownNormalize();

  function escapeMarkdownTableCell(text: string): string {
    return markdownNormalizeHelper.escapeMarkdownPipes(
      markdownNormalizeHelper.normalizeMarkdownText(text)
    );
  }

  const markdownTableEscapeApi = {
    escapeMarkdownTableCell
  };

  moduleRegistry.registerModule("markdownTableEscape", markdownTableEscapeApi);
})();
