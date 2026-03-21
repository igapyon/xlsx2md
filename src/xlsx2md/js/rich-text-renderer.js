(() => {
    const moduleRegistry = getXlsx2mdModuleRegistry();
    function createRichTextRendererApi(deps = {}) {
        const richTextParser = requireXlsx2mdRichTextParserModule().createRichTextParserApi({
            normalizeMarkdownText: deps.normalizeMarkdownText
        });
        const plainFormatter = requireXlsx2mdRichTextPlainFormatterModule().createRichTextPlainFormatterApi();
        const githubFormatter = requireXlsx2mdRichTextGithubFormatterModule().createRichTextGithubFormatterApi();
        function normalizeGithubSegment(text) {
            return githubFormatter.renderGithubTokens(richTextParser.splitTextWithLineBreaks(text));
        }
        function normalizeGithubCellText(text) {
            return normalizeGithubSegment(text)
                .replace(/ {2,}/g, " ")
                .trim();
        }
        function renderTokens(tokens, formattingMode) {
            if (!tokens.length)
                return "";
            if (formattingMode !== "github") {
                return plainFormatter.renderPlainTokens(tokens);
            }
            return githubFormatter.renderGithubTokens(tokens);
        }
        function tokenizeCellDisplayText(cell, formattingMode = "plain") {
            return richTextParser.tokenizeCellDisplayText(cell, formattingMode);
        }
        function renderCellDisplayText(cell, formattingMode = "plain") {
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
