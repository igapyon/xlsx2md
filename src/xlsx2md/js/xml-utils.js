(() => {
    const moduleRegistry = getXlsx2mdModuleRegistry();
    const textDecoder = new TextDecoder("utf-8");
    const runtimeEnv = requireXlsx2mdRuntimeEnv();
    function xmlToDocument(xmlText) {
        return runtimeEnv.xmlToDocument(xmlText);
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
            if (node.nodeType === runtimeEnv.ELEMENT_NODE && node.localName === localName) {
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
    const xmlUtilsApi = {
        xmlToDocument,
        getElementsByLocalName,
        getFirstChildByLocalName,
        getDirectChildByLocalName,
        decodeXmlText,
        getTextContent
    };
    moduleRegistry.registerModule("xmlUtils", xmlUtilsApi);
})();
