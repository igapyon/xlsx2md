(() => {
    const moduleRegistry = getXlsx2mdModuleRegistry();
    const textDecoder = new TextDecoder("utf-8");
    const runtimeEnv = requireXlsx2mdRuntimeEnv();
    function decodeXmlText(bytes) {
        return textDecoder.decode(bytes);
    }
    function xmlToDocument(xmlText) {
        return runtimeEnv.xmlToDocument(xmlText);
    }
    function getTextContent(node) {
        return ((node === null || node === void 0 ? void 0 : node.textContent) || "").replace(/\r\n/g, "\n");
    }
    function hasEnabledBooleanValue(node) {
        if (!node)
            return false;
        const value = (node.getAttribute("val") || "").trim().toLowerCase();
        return value !== "false" && value !== "0" && value !== "none";
    }
    function parseRichTextRuns(item) {
        const runElements = Array.from(item.childNodes).filter((node) => (node.nodeType === runtimeEnv.ELEMENT_NODE && node.localName === "r"));
        if (runElements.length === 0) {
            return null;
        }
        const runs = [];
        for (const runElement of runElements) {
            const text = Array.from(runElement.getElementsByTagName("t")).map((node) => getTextContent(node)).join("");
            if (!text)
                continue;
            const properties = runElement.getElementsByTagName("rPr")[0] || null;
            const run = {
                text,
                bold: hasEnabledBooleanValue(properties === null || properties === void 0 ? void 0 : properties.getElementsByTagName("b")[0]),
                italic: hasEnabledBooleanValue(properties === null || properties === void 0 ? void 0 : properties.getElementsByTagName("i")[0]),
                strike: hasEnabledBooleanValue(properties === null || properties === void 0 ? void 0 : properties.getElementsByTagName("strike")[0]),
                underline: hasEnabledBooleanValue(properties === null || properties === void 0 ? void 0 : properties.getElementsByTagName("u")[0])
            };
            const previous = runs[runs.length - 1];
            if (previous
                && previous.bold === run.bold
                && previous.italic === run.italic
                && previous.strike === run.strike
                && previous.underline === run.underline) {
                previous.text += run.text;
            }
            else {
                runs.push(run);
            }
        }
        return runs.length > 0 && runs.some((run) => run.bold || run.italic || run.strike || run.underline) ? runs : null;
    }
    function parseSharedStringEntry(item) {
        const runs = parseRichTextRuns(item);
        if (runs) {
            return {
                text: runs.map((run) => run.text).join(""),
                runs
            };
        }
        const parts = [];
        const walk = (node) => {
            if (node.nodeType === runtimeEnv.ELEMENT_NODE) {
                const element = node;
                if (element.localName === "rPh" || element.localName === "phoneticPr") {
                    return;
                }
                if (element.localName === "t") {
                    parts.push(getTextContent(element));
                    return;
                }
            }
            for (const child of Array.from(node.childNodes)) {
                walk(child);
            }
        };
        walk(item);
        return {
            text: parts.join(""),
            runs: null
        };
    }
    function parseSharedStrings(files) {
        const sharedStringsBytes = files.get("xl/sharedStrings.xml");
        if (!sharedStringsBytes) {
            return [];
        }
        const doc = xmlToDocument(decodeXmlText(sharedStringsBytes));
        const items = Array.from(doc.getElementsByTagName("si"));
        return items.map((item) => parseSharedStringEntry(item));
    }
    const sharedStringsApi = {
        parseSharedStringEntry,
        parseSharedStrings
    };
    moduleRegistry.registerModule("sharedStrings", sharedStringsApi);
})();
