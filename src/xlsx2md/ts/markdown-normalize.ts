(() => {
  const MARKDOWN_UNSAFE_UNICODE_REGEX = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F-\u009F\u00AD\u200B-\u200F\u2028\u2029\u202A-\u202E\u2060-\u206F\uFEFF\uFDD0-\uFDEF\uFFFE\uFFFF]/g;

  function normalizeMarkdownText(text: string): string {
    return String(text || "")
      .replace(MARKDOWN_UNSAFE_UNICODE_REGEX, " ")
      .replace(/\r\n?|\n/g, " ")
      .replace(/\t/g, " ");
  }

  function escapeMarkdownPipes(text: string): string {
    return String(text || "").replace(/\|/g, "\\|");
  }

  function normalizeMarkdownTableCell(text: string): string {
    return escapeMarkdownPipes(normalizeMarkdownText(text));
  }

  function normalizeMarkdownHeadingText(text: string): string {
    return normalizeMarkdownText(text).replace(/^#+\s*/, "");
  }

  function normalizeMarkdownListItemText(text: string): string {
    return normalizeMarkdownText(text).replace(/^([-*+]|\d+\.)\s+/, "");
  }

  (globalThis as typeof globalThis & {
    __xlsx2mdMarkdownNormalize?: {
      normalizeMarkdownText: typeof normalizeMarkdownText;
      escapeMarkdownPipes: typeof escapeMarkdownPipes;
      normalizeMarkdownTableCell: typeof normalizeMarkdownTableCell;
      normalizeMarkdownHeadingText: typeof normalizeMarkdownHeadingText;
      normalizeMarkdownListItemText: typeof normalizeMarkdownListItemText;
    };
  }).__xlsx2mdMarkdownNormalize = {
    normalizeMarkdownText,
    escapeMarkdownPipes,
    normalizeMarkdownTableCell,
    normalizeMarkdownHeadingText,
    normalizeMarkdownListItemText
  };
})();
