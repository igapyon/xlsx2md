(() => {
    const textDecoder = new TextDecoder("utf-8");
    function xmlToDocument(xmlText) {
        return new DOMParser().parseFromString(xmlText, "application/xml");
    }
    function getElementsByLocalName(root, localName) {
        const elements = Array.from(root.getElementsByTagName("*"));
        return elements.filter((element) => element.localName === localName);
    }
    function getFirstChildByLocalName(root, localName) {
        return getElementsByLocalName(root, localName)[0] || null;
    }
    function getDirectChildByLocalName(root, localName) {
        if (!root)
            return null;
        for (const node of Array.from(root.childNodes)) {
            if (node.nodeType === Node.ELEMENT_NODE && node.localName === localName) {
                return node;
            }
        }
        return null;
    }
    function decodeXmlText(bytes) {
        return textDecoder.decode(bytes);
    }
    function getTextContent(node) {
        return ((node === null || node === void 0 ? void 0 : node.textContent) || "").replace(/\r\n/g, "\n");
    }
    globalThis.__xlsx2mdXmlUtils = {
        xmlToDocument,
        getElementsByLocalName,
        getFirstChildByLocalName,
        getDirectChildByLocalName,
        decodeXmlText,
        getTextContent
    };
})();
