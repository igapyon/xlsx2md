(() => {
    const moduleRegistry = getXlsx2mdModuleRegistry();
    const ELEMENT_NODE = 1;
    const TEXT_NODE = 3;
    function createDomParser() {
        if (typeof DOMParser !== "function") {
            throw new Error("This environment does not provide DOMParser.");
        }
        return new DOMParser();
    }
    function xmlToDocument(xmlText) {
        return createDomParser().parseFromString(xmlText, "application/xml");
    }
    const runtimeEnvApi = {
        ELEMENT_NODE,
        TEXT_NODE,
        xmlToDocument
    };
    moduleRegistry.registerModule("runtimeEnv", runtimeEnvApi);
})();
