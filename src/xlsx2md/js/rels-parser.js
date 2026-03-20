(() => {
    const moduleRegistry = getXlsx2mdModuleRegistry();
    function createRelsParserApi(deps) {
        function normalizeZipPath(baseFilePath, targetPath) {
            const baseDirParts = baseFilePath.split("/").slice(0, -1);
            const inputParts = targetPath.split("/");
            const parts = targetPath.startsWith("/") ? [] : baseDirParts;
            for (const part of inputParts) {
                if (!part || part === ".")
                    continue;
                if (part === "..") {
                    parts.pop();
                }
                else {
                    parts.push(part);
                }
            }
            return parts.join("/");
        }
        function parseRelationships(files, relsPath, sourcePath) {
            const relBytes = files.get(relsPath);
            const relations = new Map();
            if (!relBytes) {
                return relations;
            }
            const doc = deps.xmlToDocument(deps.decodeXmlText(relBytes));
            const nodes = Array.from(doc.getElementsByTagName("Relationship"));
            for (const node of nodes) {
                const id = node.getAttribute("Id") || "";
                const target = node.getAttribute("Target") || "";
                if (!id || !target)
                    continue;
                relations.set(id, normalizeZipPath(sourcePath, target));
            }
            return relations;
        }
        function buildRelsPath(sourcePath) {
            const parts = sourcePath.split("/");
            const fileName = parts.pop() || "";
            const dir = parts.join("/");
            return `${dir}/_rels/${fileName}.rels`;
        }
        return {
            normalizeZipPath,
            parseRelationships,
            buildRelsPath
        };
    }
    const relsParserApi = {
        createRelsParserApi
    };
    moduleRegistry.registerModule("relsParser", relsParserApi);
})();
