(() => {
    const textDecoder = new TextDecoder("utf-8");
    function decodeXmlText(bytes) {
        return textDecoder.decode(bytes);
    }
    function xmlToDocument(xmlText) {
        return new DOMParser().parseFromString(xmlText, "application/xml");
    }
    function getElementsByLocalName(root, localName) {
        const elements = Array.from(root.getElementsByTagName("*"));
        return elements.filter((element) => element.localName === localName);
    }
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
        const doc = xmlToDocument(decodeXmlText(relBytes));
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
    function normalizeFormulaAddress(address) {
        return String(address || "").trim().replace(/\$/g, "").toUpperCase();
    }
    function parseRangeAddress(rawRange) {
        const match = String(rawRange || "").trim().match(/^(\$?[A-Z]+\$?\d+):(\$?[A-Z]+\$?\d+)$/i);
        if (!match)
            return null;
        return {
            start: normalizeFormulaAddress(match[1]),
            end: normalizeFormulaAddress(match[2])
        };
    }
    function normalizeStructuredTableKey(value) {
        return String(value || "").normalize("NFKC").trim().toUpperCase();
    }
    function parseWorksheetTables(files, worksheetDoc, sheetName, sheetPath) {
        const sheetRels = parseRelationships(files, buildRelsPath(sheetPath), sheetPath);
        const tablePartElements = getElementsByLocalName(worksheetDoc, "tablePart");
        const tables = [];
        for (const tablePartElement of tablePartElements) {
            const relId = tablePartElement.getAttribute("r:id") || tablePartElement.getAttribute("id") || "";
            if (!relId)
                continue;
            const tablePath = sheetRels.get(relId) || "";
            if (!tablePath)
                continue;
            const tableBytes = files.get(tablePath);
            if (!tableBytes)
                continue;
            const tableDoc = xmlToDocument(decodeXmlText(tableBytes));
            const tableElement = getElementsByLocalName(tableDoc, "table")[0] || null;
            if (!tableElement)
                continue;
            const ref = tableElement.getAttribute("ref") || "";
            const range = parseRangeAddress(ref);
            if (!range)
                continue;
            const columns = getElementsByLocalName(tableElement, "tableColumn")
                .map((columnElement) => String(columnElement.getAttribute("name") || "").trim())
                .filter(Boolean);
            tables.push({
                sheetName,
                name: tableElement.getAttribute("name") || "",
                displayName: tableElement.getAttribute("displayName") || tableElement.getAttribute("name") || "",
                start: range.start,
                end: range.end,
                columns,
                headerRowCount: Number(tableElement.getAttribute("headerRowCount") || 1) || 1,
                totalsRowCount: Number(tableElement.getAttribute("totalsRowCount") || 0) || 0
            });
        }
        return tables;
    }
    globalThis.__xlsx2mdWorksheetTables = {
        normalizeStructuredTableKey,
        parseWorksheetTables
    };
})();
