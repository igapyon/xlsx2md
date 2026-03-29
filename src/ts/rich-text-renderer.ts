/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */

(() => {
  const moduleRegistry = getXlsx2mdModuleRegistry();

  type RichTextStyle = {
    bold: boolean;
    italic: boolean;
    strike: boolean;
    underline: boolean;
  };

  type RichTextRun = RichTextStyle & {
    text: string;
  };

  type ParsedCellLike = {
    outputValue: string;
    textStyle: RichTextStyle;
    richTextRuns: RichTextRun[] | null;
  };

  type RichTextRendererDeps = {
    normalizeMarkdownText?: (text: string) => string;
  };

  type RichTextToken =
    | {
      kind: "text";
      text: string;
    }
    | {
      kind: "lineBreak";
    }
    | {
      kind: "styledText";
      parts: Array<{
        kind: "text" | "escaped";
        text: string;
        rawText: string;
      }>;
      style: RichTextStyle;
    };

  function createRichTextRendererApi(deps: RichTextRendererDeps = {}) {
    const richTextParser = requireXlsx2mdRichTextParserModule<ParsedCellLike>().createRichTextParserApi({
      normalizeMarkdownText: deps.normalizeMarkdownText
    });
    const plainFormatter = requireXlsx2mdRichTextPlainFormatterModule().createRichTextPlainFormatterApi();
    const githubFormatter = requireXlsx2mdRichTextGithubFormatterModule().createRichTextGithubFormatterApi();

    function normalizeGithubSegment(text: string): string {
      return githubFormatter.renderGithubTokens(richTextParser.splitTextWithLineBreaks(text));
    }

    function normalizeGithubCellText(text: string): string {
      return normalizeGithubSegment(text)
        .replace(/ {2,}/g, " ")
        .trim();
    }

    function renderTokens(tokens: RichTextToken[], formattingMode: "plain" | "github"): string {
      if (!tokens.length) return "";
      if (formattingMode !== "github") {
        return plainFormatter.renderPlainTokens(tokens);
      }
      return githubFormatter.renderGithubTokens(tokens);
    }

    function tokenizeCellDisplayText(
      cell: ParsedCellLike | undefined,
      formattingMode: "plain" | "github" = "plain"
    ): RichTextToken[] {
      return richTextParser.tokenizeCellDisplayText(cell, formattingMode);
    }

    function renderCellDisplayText(
      cell: ParsedCellLike | undefined,
      formattingMode: "plain" | "github" = "plain"
    ): string {
      return renderTokens(tokenizeCellDisplayText(cell, formattingMode), formattingMode);
    }

    return {
      compactText: richTextParser.compactText,
      normalizeGithubSegment,
      normalizeGithubCellText,
      applyTextStyle: githubFormatter.applyTextStyle,
      renderStyledTextParts: plainFormatter.renderStyledTextParts,
      splitTextWithLineBreaks: richTextParser.splitTextWithLineBreaks,
      tokenizePlainCellText: richTextParser.tokenizePlainCellText,
      tokenizeGithubCellText: richTextParser.tokenizeGithubCellText,
      tokenizeGithubRichTextRuns: richTextParser.tokenizeGithubRichTextRuns,
      tokenizeCellDisplayText,
      renderPlainTokens: plainFormatter.renderPlainTokens,
      renderGithubTokens: githubFormatter.renderGithubTokens,
      renderTokens,
      renderCellDisplayText
    };
  }

  const richTextRendererApi = {
    createRichTextRendererApi
  };

  moduleRegistry.registerModule("richTextRenderer", richTextRendererApi);
})();
