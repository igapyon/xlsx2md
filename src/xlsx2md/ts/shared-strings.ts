(() => {
  const moduleRegistry = getXlsx2mdModuleRegistry();
  const textDecoder = new TextDecoder("utf-8");
  const runtimeEnv = requireXlsx2mdRuntimeEnv();

  function decodeXmlText(bytes: Uint8Array): string {
    return textDecoder.decode(bytes);
  }

  function xmlToDocument(xmlText: string): Document {
    return runtimeEnv.xmlToDocument(xmlText);
  }

  function getTextContent(node: Element | null | undefined): string {
    return (node?.textContent || "").replace(/\r\n/g, "\n");
  }

  function parseSharedStrings(files: Map<string, Uint8Array>): string[] {
    const sharedStringsBytes = files.get("xl/sharedStrings.xml");
    if (!sharedStringsBytes) {
      return [];
    }
    const doc = xmlToDocument(decodeXmlText(sharedStringsBytes));
    const items = Array.from(doc.getElementsByTagName("si"));
    return items.map((item) => {
      const parts: string[] = [];
      const walk = (node: Node): void => {
        if (node.nodeType === runtimeEnv.ELEMENT_NODE) {
          const element = node as Element;
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

  const sharedStringsApi = {
    parseSharedStrings
  };

  moduleRegistry.registerModule("sharedStrings", sharedStringsApi);
})();
