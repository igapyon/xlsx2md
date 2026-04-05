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

  type RichTextParserDeps = {
    normalizeMarkdownText?: (text: string) => string;
  };

  type RawLineToken =
    | {
      kind: "text";
      rawText: string;
    }
    | {
      kind: "lineBreak";
    };

  function createRichTextParserApi(deps: RichTextParserDeps = {}) {
    const markdownEscapeHelper = requireXlsx2mdMarkdownEscape();
    const normalizeInlineText = deps.normalizeMarkdownText || ((text: string) => String(text || "").replace(/\r\n?|\n/g, " ").replace(/\t/g, " "));

    function compactText(text: string): string {
      return normalizeInlineText(markdownEscapeHelper.escapeMarkdownLiteralText(text)).replace(/\s+/g, " ").trim();
    }

    function splitRawTextWithLineBreaks(text: string): RawLineToken[] {
      const normalized = String(text || "")
        .replace(/\r\n?/g, "\n")
        .replace(/\t/g, " ");
      if (!normalized) return [];
      const parts = normalized.split("\n");
      const tokens: RawLineToken[] = [];
      for (let index = 0; index < parts.length; index += 1) {
        if (parts[index]) {
          tokens.push({
            kind: "text",
            rawText: parts[index]
          });
        }
        if (index < parts.length - 1) {
          tokens.push({ kind: "lineBreak" });
        }
      }
      return tokens;
    }

    function splitTextWithLineBreaks(text: string): RichTextToken[] {
      return splitRawTextWithLineBreaks(text).map((token) => {
        if (token.kind === "lineBreak") return token;
        return {
          kind: "text",
          text: markdownEscapeHelper.escapeMarkdownLiteralText(token.rawText)
        };
      });
    }

    function createStyledTextToken(text: string, style: RichTextStyle): RichTextToken {
      return {
        kind: "styledText",
        parts: markdownEscapeHelper.escapeMarkdownLiteralParts(text),
        style
      };
    }

    function tokenizePlainCellText(text: string): RichTextToken[] {
      const compacted = compactText(text);
      if (!compacted) return [];
      return [{ kind: "text", text: compacted }];
    }

    function tokenizeGithubCellText(text: string, style: RichTextStyle): RichTextToken[] {
      const tokens = splitRawTextWithLineBreaks(text);
      if (!tokens.length) return [];
      return tokens.map((token) => {
        if (token.kind !== "text") return token;
        return createStyledTextToken(token.rawText, style);
      });
    }

    function tokenizeGithubRichTextRuns(runs: RichTextRun[]): RichTextToken[] {
      return runs.flatMap((run) => splitRawTextWithLineBreaks(run.text).map((token) => {
        if (token.kind !== "text") return token;
        return createStyledTextToken(token.rawText, {
          bold: run.bold,
          italic: run.italic,
          strike: run.strike,
          underline: run.underline
        });
      }));
    }

    function tokenizeCellDisplayText(
      cell: ParsedCellLike | undefined,
      formattingMode: "plain" | "github" = "plain"
    ): RichTextToken[] {
      if (!cell) return [];
      if (formattingMode !== "github") {
        return tokenizePlainCellText(String(cell.outputValue || ""));
      }
      const displayValue = compactText(String(cell.outputValue || ""));
      if (cell.richTextRuns && displayValue === compactText(cell.richTextRuns.map((run) => run.text).join(""))) {
        return tokenizeGithubRichTextRuns(cell.richTextRuns);
      }
      return tokenizeGithubCellText(String(cell.outputValue || ""), cell.textStyle);
    }

    return {
      compactText,
      splitRawTextWithLineBreaks,
      splitTextWithLineBreaks,
      createStyledTextToken,
      tokenizePlainCellText,
      tokenizeGithubCellText,
      tokenizeGithubRichTextRuns,
      tokenizeCellDisplayText
    };
  }

  const richTextParserApi = {
    createRichTextParserApi
  };

  moduleRegistry.registerModule("richTextParser", richTextParserApi);
})();
