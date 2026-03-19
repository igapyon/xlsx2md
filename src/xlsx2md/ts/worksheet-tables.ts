(() => {
  type ParsedTable = {
    sheetName: string;
    name: string;
    displayName: string;
    start: string;
    end: string;
    columns: string[];
    headerRowCount: number;
    totalsRowCount: number;
  };

  const textDecoder = new TextDecoder("utf-8");

  function decodeXmlText(bytes: Uint8Array): string {
    return textDecoder.decode(bytes);
  }

  function xmlToDocument(xmlText: string): Document {
    return new DOMParser().parseFromString(xmlText, "application/xml");
  }

  function getElementsByLocalName(root: ParentNode, localName: string): Element[] {
    const elements = Array.from(root.getElementsByTagName("*"));
    return elements.filter((element) => element.localName === localName);
  }

  function normalizeZipPath(baseFilePath: string, targetPath: string): string {
    const baseDirParts = baseFilePath.split("/").slice(0, -1);
    const inputParts = targetPath.split("/");
    const parts = targetPath.startsWith("/") ? [] : baseDirParts;
    for (const part of inputParts) {
      if (!part || part === ".") continue;
      if (part === "..") {
        parts.pop();
      } else {
        parts.push(part);
      }
    }
    return parts.join("/");
  }

  function parseRelationships(files: Map<string, Uint8Array>, relsPath: string, sourcePath: string): Map<string, string> {
    const relBytes = files.get(relsPath);
    const relations = new Map<string, string>();
    if (!relBytes) {
      return relations;
    }
    const doc = xmlToDocument(decodeXmlText(relBytes));
    const nodes = Array.from(doc.getElementsByTagName("Relationship"));
    for (const node of nodes) {
      const id = node.getAttribute("Id") || "";
      const target = node.getAttribute("Target") || "";
      if (!id || !target) continue;
      relations.set(id, normalizeZipPath(sourcePath, target));
    }
    return relations;
  }

  function buildRelsPath(sourcePath: string): string {
    const parts = sourcePath.split("/");
    const fileName = parts.pop() || "";
    const dir = parts.join("/");
    return `${dir}/_rels/${fileName}.rels`;
  }

  function normalizeFormulaAddress(address: string): string {
    return String(address || "").trim().replace(/\$/g, "").toUpperCase();
  }

  function parseRangeAddress(rawRange: string): { start: string; end: string } | null {
    const match = String(rawRange || "").trim().match(/^(\$?[A-Z]+\$?\d+):(\$?[A-Z]+\$?\d+)$/i);
    if (!match) return null;
    return {
      start: normalizeFormulaAddress(match[1]),
      end: normalizeFormulaAddress(match[2])
    };
  }

  function normalizeStructuredTableKey(value: string): string {
    return String(value || "").normalize("NFKC").trim().toUpperCase();
  }

  function parseWorksheetTables(
    files: Map<string, Uint8Array>,
    worksheetDoc: Document,
    sheetName: string,
    sheetPath: string
  ): ParsedTable[] {
    const sheetRels = parseRelationships(files, buildRelsPath(sheetPath), sheetPath);
    const tablePartElements = getElementsByLocalName(worksheetDoc, "tablePart");
    const tables: ParsedTable[] = [];

    for (const tablePartElement of tablePartElements) {
      const relId = tablePartElement.getAttribute("r:id") || tablePartElement.getAttribute("id") || "";
      if (!relId) continue;
      const tablePath = sheetRels.get(relId) || "";
      if (!tablePath) continue;
      const tableBytes = files.get(tablePath);
      if (!tableBytes) continue;
      const tableDoc = xmlToDocument(decodeXmlText(tableBytes));
      const tableElement = getElementsByLocalName(tableDoc, "table")[0] || null;
      if (!tableElement) continue;
      const ref = tableElement.getAttribute("ref") || "";
      const range = parseRangeAddress(ref);
      if (!range) continue;
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

  (globalThis as typeof globalThis & {
    __xlsx2mdWorksheetTables?: {
      normalizeStructuredTableKey: typeof normalizeStructuredTableKey;
      parseWorksheetTables: typeof parseWorksheetTables;
    };
  }).__xlsx2mdWorksheetTables = {
    normalizeStructuredTableKey,
    parseWorksheetTables
  };
})();
