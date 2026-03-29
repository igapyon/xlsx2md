/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */

(() => {
  const moduleRegistry = getXlsx2mdModuleRegistry();

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
      style: {
        bold: boolean;
        italic: boolean;
        strike: boolean;
        underline: boolean;
      };
    };

  function createRichTextPlainFormatterApi() {
    function renderStyledTextPart(part: { kind: "text" | "escaped"; text: string; rawText: string }): string {
      if (part.kind === "escaped") {
        return part.text;
      }
      return part.text;
    }

    function renderStyledTextParts(parts: Array<{ kind: "text" | "escaped"; text: string; rawText: string }>): string {
      return parts.map((part) => renderStyledTextPart(part)).join("");
    }

    function renderPlainTokens(tokens: RichTextToken[]): string {
      if (!tokens.length) return "";
      return tokens
        .map((token) => {
          if (token.kind === "lineBreak") return " ";
          if (token.kind === "styledText") return renderStyledTextParts(token.parts);
          return token.text;
        })
        .join("")
        .replace(/ {2,}/g, " ")
        .trim();
    }

    return {
      renderStyledTextPart,
      renderStyledTextParts,
      renderPlainTokens
    };
  }

  const richTextPlainFormatterApi = {
    createRichTextPlainFormatterApi
  };

  moduleRegistry.registerModule("richTextPlainFormatter", richTextPlainFormatterApi);
})();
