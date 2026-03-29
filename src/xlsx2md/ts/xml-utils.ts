/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */

(() => {
  const moduleRegistry = getXlsx2mdModuleRegistry();
  const textDecoder = new TextDecoder("utf-8");
  const runtimeEnv = requireXlsx2mdRuntimeEnv();

  function xmlToDocument(xmlText: string): Document {
    return runtimeEnv.xmlToDocument(xmlText);
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
      if (node.nodeType === runtimeEnv.ELEMENT_NODE && (node as Element).localName === localName) {
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
