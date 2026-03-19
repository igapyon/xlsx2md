(() => {
    const textDecoder = new TextDecoder("utf-8");
    function decodeXmlText(bytes) {
        return textDecoder.decode(bytes);
    }
    function xmlToDocument(xmlText) {
        return new DOMParser().parseFromString(xmlText, "application/xml");
    }
    function getTextContent(node) {
        return ((node === null || node === void 0 ? void 0 : node.textContent) || "").replace(/\r\n/g, "\n");
    }
    function parseSharedStrings(files) {
        const sharedStringsBytes = files.get("xl/sharedStrings.xml");
        if (!sharedStringsBytes) {
            return [];
        }
        const doc = xmlToDocument(decodeXmlText(sharedStringsBytes));
        const items = Array.from(doc.getElementsByTagName("si"));
        return items.map((item) => {
            const parts = [];
            const walk = (node) => {
                if (node.nodeType === Node.ELEMENT_NODE) {
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
            return parts.join("");
        });
    }
    globalThis.__xlsx2mdSharedStrings = {
        parseSharedStrings
    };
})();
