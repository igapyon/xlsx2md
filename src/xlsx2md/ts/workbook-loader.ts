(() => {
  const moduleRegistry = getXlsx2mdModuleRegistry();
  type ParsedWorkbook = {
    name: string;
    sheets: unknown[];
    sharedStrings: string[];
    definedNames: {
      name: string;
      formulaText: string;
      localSheetName: string | null;
    }[];
  };

  type CellStyleInfo = {
    borders: unknown;
    numFmtId: number;
    formatCode: string;
  };

  type WorkbookLoaderDependencies = {
    unzipEntries: (arrayBuffer: ArrayBuffer) => Promise<Map<string, Uint8Array>>;
    parseSharedStrings: (files: Map<string, Uint8Array>) => string[];
    parseCellStyles: (files: Map<string, Uint8Array>) => CellStyleInfo[];
    parseRelationships: (files: Map<string, Uint8Array>, relsPath: string, sourcePath: string) => Map<string, string>;
    xmlToDocument: (xmlText: string) => Document;
    decodeXmlText: (bytes: Uint8Array) => string;
    getTextContent: (node: Element | null | undefined) => string;
    parseWorksheet: (
      files: Map<string, Uint8Array>,
      sheetName: string,
      sheetPath: string,
      sheetIndex: number,
      sharedStrings: string[],
      cellStyles: CellStyleInfo[]
    ) => unknown;
    postProcessWorkbook?: (workbook: ParsedWorkbook) => void;
  };

  function parseDefinedNames(
    workbookDoc: Document,
    sheetNames: string[],
    getTextContent: WorkbookLoaderDependencies["getTextContent"]
  ): {
    name: string;
    formulaText: string;
    localSheetName: string | null;
  }[] {
    const result: {
      name: string;
      formulaText: string;
      localSheetName: string | null;
    }[] = [];
    const definedNameElements = Array.from(workbookDoc.getElementsByTagName("definedName"));
    for (const element of definedNameElements) {
      const name = element.getAttribute("name") || "";
      if (!name || name.startsWith("_xlnm.")) continue;
      const formulaText = getTextContent(element).trim();
      if (!formulaText) continue;
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

  async function parseWorkbook(
    arrayBuffer: ArrayBuffer,
    workbookName: string,
    deps: WorkbookLoaderDependencies
  ): Promise<ParsedWorkbook> {
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
    const workbook: ParsedWorkbook = {
      name: workbookName,
      sheets,
      sharedStrings,
      definedNames
    };
    deps.postProcessWorkbook?.(workbook);
    return workbook;
  }

  const workbookLoaderApi = {
    parseDefinedNames,
    parseWorkbook
  };

  moduleRegistry.registerModule("workbookLoader", workbookLoaderApi);
})();
