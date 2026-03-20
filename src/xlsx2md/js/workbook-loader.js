(() => {
    const moduleRegistry = getXlsx2mdModuleRegistry();
    function parseDefinedNames(workbookDoc, sheetNames, getTextContent) {
        const result = [];
        const definedNameElements = Array.from(workbookDoc.getElementsByTagName("definedName"));
        for (const element of definedNameElements) {
            const name = element.getAttribute("name") || "";
            if (!name || name.startsWith("_xlnm."))
                continue;
            const formulaText = getTextContent(element).trim();
            if (!formulaText)
                continue;
            const localSheetIdText = element.getAttribute("localSheetId");
            const localSheetId = localSheetIdText == null || localSheetIdText === "" ? Number.NaN : Number(localSheetIdText);
            result.push({
                name,
                formulaText: formulaText.startsWith("=") ? formulaText : `=${formulaText}`,
                localSheetName: Number.isNaN(localSheetId) ? null : (sheetNames[localSheetId] || null)
            });
        }
        return result;
    }
    async function parseWorkbook(arrayBuffer, workbookName, deps) {
        var _a;
        const files = await deps.unzipEntries(arrayBuffer);
        const workbookBytes = files.get("xl/workbook.xml");
        if (!workbookBytes) {
            throw new Error("xl/workbook.xml was not found.");
        }
        const sharedStrings = deps.parseSharedStrings(files);
        const cellStyles = deps.parseCellStyles(files);
        const rels = deps.parseRelationships(files, "xl/_rels/workbook.xml.rels", "xl/workbook.xml");
        const workbookDoc = deps.xmlToDocument(deps.decodeXmlText(workbookBytes));
        const sheetNodes = Array.from(workbookDoc.getElementsByTagName("sheet"));
        const sheetNames = sheetNodes.map((sheetNode, index) => sheetNode.getAttribute("name") || `Sheet${index + 1}`);
        const definedNames = parseDefinedNames(workbookDoc, sheetNames, deps.getTextContent);
        const sheets = sheetNodes.map((sheetNode, index) => {
            const name = sheetNode.getAttribute("name") || `Sheet${index + 1}`;
            const relId = sheetNode.getAttribute("r:id") || "";
            const sheetPath = rels.get(relId) || "";
            return deps.parseWorksheet(files, name, sheetPath, index + 1, sharedStrings, cellStyles);
        });
        const workbook = {
            name: workbookName,
            sheets,
            sharedStrings,
            definedNames
        };
        (_a = deps.postProcessWorkbook) === null || _a === void 0 ? void 0 : _a.call(deps, workbook);
        return workbook;
    }
    const workbookLoaderApi = {
        parseDefinedNames,
        parseWorkbook
    };
    moduleRegistry.registerModule("workbookLoader", workbookLoaderApi);
})();
