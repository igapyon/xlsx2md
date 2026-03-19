(() => {
  const textDecoder = new TextDecoder("utf-8");

  function decodeXmlText(bytes: Uint8Array): string {
    return textDecoder.decode(bytes);
  }

  function xmlToDocument(xmlText: string): Document {
    return new DOMParser().parseFromString(xmlText, "application/xml");
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
        if (node.nodeType === Node.ELEMENT_NODE) {
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

  (globalThis as typeof globalThis & {
    __xlsx2mdSharedStrings?: {
      parseSharedStrings: typeof parseSharedStrings;
    };
  }).__xlsx2mdSharedStrings = {
    parseSharedStrings
  };
})();
