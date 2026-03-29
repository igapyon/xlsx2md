/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */

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
