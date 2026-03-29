/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const moduleRegistry = getXlsx2mdModuleRegistry();
    function createRichTextParserApi(deps = {}) {
        const markdownEscapeHelper = requireXlsx2mdMarkdownEscape();
        const normalizeInlineText = deps.normalizeMarkdownText || ((text) => String(text || "").replace(/\r\n?|\n/g, " ").replace(/\t/g, " "));
        function compactText(text) {
            return normalizeInlineText(markdownEscapeHelper.escapeMarkdownLiteralText(text)).replace(/\s+/g, " ").trim();
        }
        function splitRawTextWithLineBreaks(text) {
            const normalized = String(text || "")
                .replace(/\r\n?/g, "\n")
                .replace(/\t/g, " ");
            if (!normalized)
                return [];
            const parts = normalized.split("\n");
            const tokens = [];
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
        function splitTextWithLineBreaks(text) {
            return splitRawTextWithLineBreaks(text).map((token) => {
                if (token.kind === "lineBreak")
                    return token;
                return {
                    kind: "text",
                    text: markdownEscapeHelper.escapeMarkdownLiteralText(token.rawText)
                };
            });
        }
        function createStyledTextToken(text, style) {
            return {
                kind: "styledText",
                parts: markdownEscapeHelper.escapeMarkdownLiteralParts(text),
                style
            };
        }
        function tokenizePlainCellText(text) {
            const compacted = compactText(text);
            if (!compacted)
                return [];
            return [{ kind: "text", text: compacted }];
        }
        function tokenizeGithubCellText(text, style) {
            const tokens = splitRawTextWithLineBreaks(text);
            if (!tokens.length)
                return [];
            return tokens.map((token) => {
                if (token.kind !== "text")
                    return token;
                return createStyledTextToken(token.rawText, style);
            });
        }
        function tokenizeGithubRichTextRuns(runs) {
            return runs.flatMap((run) => splitRawTextWithLineBreaks(run.text).map((token) => {
                if (token.kind !== "text")
                    return token;
                return createStyledTextToken(token.rawText, {
                    bold: run.bold,
                    italic: run.italic,
                    strike: run.strike,
                    underline: run.underline
                });
            }));
        }
        function tokenizeCellDisplayText(cell, formattingMode = "plain") {
            if (!cell)
                return [];
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
