(() => {
    const moduleRegistry = getXlsx2mdModuleRegistry();
    function escapeMarkdownLineStart(text) {
        return String(text || "")
            .replace(/^(\s*)([#>])/u, "$1\\$2")
            .replace(/^(\s*)([-+*])(\s+)/u, "$1\\$2$3")
            .replace(/^(\s*)(\d+)\.(\s+)/u, "$1$2\\.$3");
    }
    function escapeMarkdownLiteralParts(text) {
        const source = String(text || "");
        const parts = [];
        let buffer = "";
        function pushTextBuffer() {
            if (!buffer)
                return;
            parts.push({ kind: "text", text: buffer, rawText: buffer });
            buffer = "";
        }
        function pushEscaped(textValue, rawText) {
            pushTextBuffer();
            if (!textValue)
                return;
            parts.push({ kind: "escaped", text: textValue, rawText });
        }
        for (let index = 0; index < source.length; index += 1) {
            const ch = source[index];
            const atLineStart = index === 0;
            const next = source[index + 1] || "";
            if (ch === "\\") {
                pushEscaped("\\\\", ch);
                continue;
            }
            if (ch === "&") {
                pushEscaped("&amp;", ch);
                continue;
            }
            if (ch === "<") {
                pushEscaped("&lt;", ch);
                continue;
            }
            if (ch === ">") {
                if (atLineStart) {
                    pushEscaped("&gt;", ch);
                    continue;
                }
                pushEscaped("&gt;", ch);
                continue;
            }
            if (/[`*_{}\[\]()!|~]/.test(ch)) {
                pushEscaped(`\\${ch}`, ch);
                continue;
            }
            if (atLineStart && /[#]/.test(ch)) {
                pushEscaped(`\\${ch}`, ch);
                continue;
            }
            if (atLineStart && /[-+*]/.test(ch) && /\s/u.test(next)) {
                pushEscaped(`\\${ch}`, ch);
                continue;
            }
            if (atLineStart && /\d/u.test(ch)) {
                let digitRun = ch;
                let cursor = index + 1;
                while (cursor < source.length && /\d/u.test(source[cursor])) {
                    digitRun += source[cursor];
                    cursor += 1;
                }
                if (source[cursor] === "." && /\s/u.test(source[cursor + 1] || "")) {
                    pushTextBuffer();
                    parts.push({ kind: "text", text: digitRun, rawText: digitRun });
                    parts.push({ kind: "escaped", text: "\\.", rawText: "." });
                    index = cursor;
                    continue;
                }
            }
            buffer += ch;
        }
        pushTextBuffer();
        return parts;
    }
    function escapeMarkdownLiteralText(text) {
        return String(text || "")
            .replace(/\r\n?/g, "\n")
            .split("\n")
            .map((line) => escapeMarkdownLiteralParts(line).map((part) => part.text).join(""))
            .join("\n");
    }
    const markdownEscapeApi = {
        escapeMarkdownLineStart,
        escapeMarkdownLiteralParts,
        escapeMarkdownLiteralText
    };
    moduleRegistry.registerModule("markdownEscape", markdownEscapeApi);
})();
