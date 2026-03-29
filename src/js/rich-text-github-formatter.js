/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const moduleRegistry = getXlsx2mdModuleRegistry();
    function createRichTextGithubFormatterApi() {
        function applyTextStyle(text, style) {
            if (!text)
                return "";
            let result = text;
            if (style.underline)
                result = `<ins>${result}</ins>`;
            if (style.strike)
                result = `~~${result}~~`;
            if (style.italic)
                result = `*${result}*`;
            if (style.bold)
                result = `**${result}**`;
            return result;
        }
        function renderStyledTextPart(part) {
            if (part.kind === "escaped") {
                return part.text;
            }
            return part.text;
        }
        function renderStyledTextParts(parts) {
            return parts.map((part) => renderStyledTextPart(part)).join("");
        }
        function renderGithubTokens(tokens) {
            if (!tokens.length)
                return "";
            return tokens
                .map((token) => {
                if (token.kind === "lineBreak")
                    return "<br>";
                if (token.kind === "styledText")
                    return applyTextStyle(renderStyledTextParts(token.parts), token.style);
                return token.text;
            })
                .join("")
                .replace(/ {2,}/g, " ")
                .trim();
        }
        return {
            applyTextStyle,
            renderStyledTextPart,
            renderStyledTextParts,
            renderGithubTokens
        };
    }
    const richTextGithubFormatterApi = {
        createRichTextGithubFormatterApi
    };
    moduleRegistry.registerModule("richTextGithubFormatter", richTextGithubFormatterApi);
})();
