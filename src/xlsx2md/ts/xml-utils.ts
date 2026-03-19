(() => {
  const textDecoder = new TextDecoder("utf-8");

  function xmlToDocument(xmlText: string): Document {
    return new DOMParser().parseFromString(xmlText, "application/xml");
  }

  function getElementsByLocalName(root: ParentNode, localName: string): Element[] {
    const elements = Array.from(root.getElementsByTagName("*"));
    return elements.filter((element) => element.localName === localName);
  }

  function getFirstChildByLocalName(root: ParentNode, localName: string): Element | null {
    return getElementsByLocalName(root, localName)[0] || null;
  }

  function getDirectChildByLocalName(root: ParentNode | null, localName: string): Element | null {
    if (!root) return null;
    for (const node of Array.from(root.childNodes)) {
      if (node.nodeType === Node.ELEMENT_NODE && (node as Element).localName === localName) {
        return node as Element;
      }
    }
    return null;
  }

  function decodeXmlText(bytes: Uint8Array): string {
    return textDecoder.decode(bytes);
  }

  function getTextContent(node: Element | null | undefined): string {
    return (node?.textContent || "").replace(/\r\n/g, "\n");
  }

  (globalThis as typeof globalThis & {
    __xlsx2mdXmlUtils?: {
      xmlToDocument: typeof xmlToDocument;
      getElementsByLocalName: typeof getElementsByLocalName;
      getFirstChildByLocalName: typeof getFirstChildByLocalName;
      getDirectChildByLocalName: typeof getDirectChildByLocalName;
      decodeXmlText: typeof decodeXmlText;
      getTextContent: typeof getTextContent;
    };
  }).__xlsx2mdXmlUtils = {
    xmlToDocument,
    getElementsByLocalName,
    getFirstChildByLocalName,
    getDirectChildByLocalName,
    decodeXmlText,
    getTextContent
  };
})();
