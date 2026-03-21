(() => {
  const moduleRegistry = getXlsx2mdModuleRegistry();
  const textDecoder = new TextDecoder("utf-8");
  const runtimeEnv = requireXlsx2mdRuntimeEnv();
  type RichTextRun = {
    text: string;
    bold: boolean;
    italic: boolean;
    strike: boolean;
    underline: boolean;
  };
  type SharedStringEntry = {
    text: string;
    runs: RichTextRun[] | null;
  };

  function decodeXmlText(bytes: Uint8Array): string {
    return textDecoder.decode(bytes);
  }

  function xmlToDocument(xmlText: string): Document {
    return runtimeEnv.xmlToDocument(xmlText);
  }

  function getTextContent(node: Element | null | undefined): string {
    return (node?.textContent || "").replace(/\r\n/g, "\n");
  }

  function hasEnabledBooleanValue(node: Element | null | undefined): boolean {
    if (!node) return false;
    const value = (node.getAttribute("val") || "").trim().toLowerCase();
    return value !== "false" && value !== "0" && value !== "none";
  }

  function parseRichTextRuns(item: Element): RichTextRun[] | null {
    const runElements = Array.from(item.childNodes).filter((node): node is Element => (
      node.nodeType === runtimeEnv.ELEMENT_NODE && (node as Element).localName === "r"
    ));
    if (runElements.length === 0) {
      return null;
    }
    const runs: RichTextRun[] = [];
    for (const runElement of runElements) {
      const text = Array.from(runElement.getElementsByTagName("t")).map((node) => getTextContent(node)).join("");
      if (!text) continue;
      const properties = runElement.getElementsByTagName("rPr")[0] || null;
      const run: RichTextRun = {
        text,
        bold: hasEnabledBooleanValue(properties?.getElementsByTagName("b")[0]),
        italic: hasEnabledBooleanValue(properties?.getElementsByTagName("i")[0]),
        strike: hasEnabledBooleanValue(properties?.getElementsByTagName("strike")[0]),
        underline: hasEnabledBooleanValue(properties?.getElementsByTagName("u")[0])
      };
      const previous = runs[runs.length - 1];
      if (
        previous
        && previous.bold === run.bold
        && previous.italic === run.italic
        && previous.strike === run.strike
        && previous.underline === run.underline
      ) {
        previous.text += run.text;
      } else {
        runs.push(run);
      }
    }
    return runs.length > 0 && runs.some((run) => run.bold || run.italic || run.strike || run.underline) ? runs : null;
  }

  function parseSharedStringEntry(item: Element): SharedStringEntry {
    const runs = parseRichTextRuns(item);
    if (runs) {
      return {
        text: runs.map((run) => run.text).join(""),
        runs
      };
    }
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
    return {
      text: parts.join(""),
      runs: null
    };
  }

  function parseSharedStrings(files: Map<string, Uint8Array>): SharedStringEntry[] {
    const sharedStringsBytes = files.get("xl/sharedStrings.xml");
    if (!sharedStringsBytes) {
      return [];
    }
    const doc = xmlToDocument(decodeXmlText(sharedStringsBytes));
    const items = Array.from(doc.getElementsByTagName("si"));
    return items.map((item) => parseSharedStringEntry(item));
  }

  const sharedStringsApi = {
    parseSharedStringEntry,
    parseSharedStrings
  };

  moduleRegistry.registerModule("sharedStrings", sharedStringsApi);
})();
