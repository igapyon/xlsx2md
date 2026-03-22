(() => {
  const moduleRegistry = getXlsx2mdModuleRegistry();
  type BorderFlags = {
    top: boolean;
    bottom: boolean;
    left: boolean;
    right: boolean;
  };

  type FormulaResolutionStatus = "resolved" | "fallback_formula" | "unsupported_external" | null;
  type FormulaResolutionSource = "cached_value" | "ast_evaluator" | "legacy_resolver" | "formula_text" | "external_unsupported" | null;
  type CachedValueState = "present_nonempty" | "present_empty" | "absent" | null;
  type RichTextStyle = {
    bold: boolean;
    italic: boolean;
    strike: boolean;
    underline: boolean;
  };
  type RichTextRun = RichTextStyle & {
    text: string;
  };
  type SharedStringEntry = {
    text: string;
    runs: RichTextRun[] | null;
  };

  type CellStyleInfo = {
    borders: BorderFlags;
    numFmtId: number;
    formatCode: string;
    textStyle: RichTextStyle;
  };

  type MergeRange = {
    startRow: number;
    startCol: number;
    endRow: number;
    endCol: number;
    ref: string;
  };

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

  type ParsedImageAsset = {
    sheetName: string;
    filename: string;
    path: string;
    anchor: string;
    data: Uint8Array;
    mediaPath: string;
  };

  type ParsedChartAsset = {
    sheetName: string;
    anchor: string;
    chartPath: string;
    title: string;
    chartType: string;
    series: {
      name: string;
      categoriesRef: string;
      valuesRef: string;
      axis: "primary" | "secondary";
    }[];
  };

  type ParsedShapeAsset = {
    sheetName: string;
    anchor: string;
    name: string;
    kind: string;
    text: string;
    widthEmu: number | null;
    heightEmu: number | null;
    elementName: string;
    anchorElementName: string;
    rawEntries: {
      key: string;
      value: string;
    }[];
    bbox: {
      left: number;
      top: number;
      right: number;
      bottom: number;
    };
    svgFilename: string | null;
    svgPath: string | null;
    svgData: Uint8Array | null;
  };

  type ParsedCell = {
    address: string;
    row: number;
    col: number;
    valueType: string;
    rawValue: string;
    outputValue: string;
    formulaText: string;
    resolutionStatus: FormulaResolutionStatus;
    resolutionSource: FormulaResolutionSource;
    cachedValueState: CachedValueState;
    styleIndex: number;
    borders: BorderFlags;
    numFmtId: number;
    formatCode: string;
    textStyle: RichTextStyle;
    richTextRuns: RichTextRun[] | null;
    formulaType: string;
    spillRef: string;
    hyperlink: {
      kind: "external" | "internal";
      target: string;
      location: string;
      tooltip: string;
      display: string;
    } | null;
  };

  type ParsedSheet = {
    name: string;
    index: number;
    path: string;
    cells: ParsedCell[];
    merges: MergeRange[];
    tables: ParsedTable[];
    images: ParsedImageAsset[];
    charts: ParsedChartAsset[];
    shapes: ParsedShapeAsset[];
    maxRow: number;
    maxCol: number;
  };

  type ExtractedCellOutput = {
    valueType: string;
    rawValue: string;
    outputValue: string;
    formulaText: string;
    resolutionStatus: FormulaResolutionStatus;
    resolutionSource: FormulaResolutionSource;
    cachedValueState: CachedValueState;
    richTextRuns: RichTextRun[] | null;
  };

  type WorksheetParserDependencies = {
    EMPTY_BORDERS: BorderFlags;
    xmlToDocument: (xmlText: string) => Document;
    decodeXmlText: (bytes: Uint8Array) => string;
    getTextContent: (node: Element | null | undefined) => string;
    parseCellAddress: (address: string) => { row: number; col: number };
    parseRangeRef: (ref: string) => MergeRange;
    parseWorksheetTables: (
      files: Map<string, Uint8Array>,
      worksheetDoc: Document,
      sheetName: string,
      sheetPath: string
    ) => ParsedTable[];
    parseDrawingImages: (
      files: Map<string, Uint8Array>,
      sheetName: string,
      sheetPath: string,
      deps: Record<string, unknown>
    ) => ParsedImageAsset[];
    parseDrawingCharts: (
      files: Map<string, Uint8Array>,
      sheetName: string,
      sheetPath: string,
      deps: Record<string, unknown>
    ) => ParsedChartAsset[];
    parseDrawingShapes: (
      files: Map<string, Uint8Array>,
      sheetName: string,
      sheetPath: string,
      deps: Record<string, unknown>
    ) => ParsedShapeAsset[];
    parseRelationshipEntries: (
      files: Map<string, Uint8Array>,
      relsPath: string,
      sourcePath: string
    ) => Map<string, {
      target: string;
      targetMode: string;
      type: string;
    }>;
    buildRelsPath: (sourcePath: string) => string;
    formatCellDisplayValue: (rawValue: string, cellStyle: CellStyleInfo) => string | null;
    buildAssetDeps: () => Record<string, unknown>;
  };

  function expandRangeAddresses(
    ref: string,
    deps: Pick<WorksheetParserDependencies, "parseRangeRef"> & { colToLetters: (col: number) => string }
  ): string[] {
    const range = deps.parseRangeRef(ref);
    const addresses: string[] = [];
    for (let row = Math.max(1, range.startRow); row <= Math.max(range.startRow, range.endRow); row += 1) {
      for (let col = Math.max(1, range.startCol); col <= Math.max(range.startCol, range.endCol); col += 1) {
        addresses.push(`${deps.colToLetters(col)}${row}`);
      }
    }
    return addresses;
  }

  function parseWorksheetHyperlinks(
    files: Map<string, Uint8Array>,
    worksheetDoc: Document,
    sheetPath: string,
    deps: Pick<WorksheetParserDependencies, "parseRelationshipEntries" | "buildRelsPath" | "parseRangeRef"> & { colToLetters: (col: number) => string }
  ): Map<string, ParsedCell["hyperlink"]> {
    const hyperlinks = new Map<string, ParsedCell["hyperlink"]>();
    const relsPath = deps.buildRelsPath(sheetPath);
    const relEntries = deps.parseRelationshipEntries(files, relsPath, sheetPath);
    const hyperlinkNodes = Array.from(worksheetDoc.getElementsByTagName("hyperlink"));
    for (const node of hyperlinkNodes) {
      const ref = (node.getAttribute("ref") || "").trim();
      if (!ref) continue;
      const relId = (node.getAttribute("r:id") || node.getAttribute("id") || "").trim();
      const relEntry = relId ? relEntries.get(relId) : null;
      const display = (node.getAttribute("display") || "").trim();
      const tooltip = (node.getAttribute("tooltip") || "").trim();
      const location = (node.getAttribute("location") || "").trim().replace(/^#/, "");
      const rawTarget = relEntry?.target?.trim() || "";
      const kind: "external" | "internal" | null = relEntry?.targetMode.toLowerCase() === "external"
        ? "external"
        : location
          ? "internal"
          : rawTarget.startsWith("#")
            ? "internal"
            : rawTarget
              ? "external"
              : null;
      if (!kind) continue;
      const target = kind === "internal"
        ? (location || rawTarget.replace(/^#/, ""))
        : rawTarget;
      if (!target) continue;
      const hyperlink = {
        kind,
        target,
        location: location || (kind === "internal" ? target : ""),
        tooltip,
        display
      } satisfies NonNullable<ParsedCell["hyperlink"]>;
      for (const address of expandRangeAddresses(ref, deps)) {
        hyperlinks.set(address, hyperlink);
      }
    }
    return hyperlinks;
  }

  function hasEnabledBooleanValue(node: Element | null | undefined): boolean {
    if (!node) return false;
    const value = (node.getAttribute("val") || "").trim().toLowerCase();
    return value !== "false" && value !== "0" && value !== "none";
  }

  function mergeTextStyle(base: RichTextStyle, override: RichTextStyle): RichTextStyle {
    return {
      bold: base.bold || override.bold,
      italic: base.italic || override.italic,
      strike: base.strike || override.strike,
      underline: base.underline || override.underline
    };
  }

  function hasTextStyle(style: RichTextStyle): boolean {
    return style.bold || style.italic || style.strike || style.underline;
  }

  function parseRichTextStyle(runProperties: Element | null | undefined): RichTextStyle {
    return {
      bold: hasEnabledBooleanValue(runProperties?.getElementsByTagName("b")[0]),
      italic: hasEnabledBooleanValue(runProperties?.getElementsByTagName("i")[0]),
      strike: hasEnabledBooleanValue(runProperties?.getElementsByTagName("strike")[0]),
      underline: hasEnabledBooleanValue(runProperties?.getElementsByTagName("u")[0])
    };
  }

  function mergeAdjacentRuns(runs: RichTextRun[]): RichTextRun[] | null {
    const merged: RichTextRun[] = [];
    for (const run of runs) {
      if (!run.text) continue;
      const previous = merged[merged.length - 1];
      if (
        previous
        && previous.bold === run.bold
        && previous.italic === run.italic
        && previous.strike === run.strike
        && previous.underline === run.underline
      ) {
        previous.text += run.text;
      } else {
        merged.push({ ...run });
      }
    }
    return merged.length > 0 && merged.some((run) => hasTextStyle(run)) ? merged : null;
  }

  function createStyledRuns(text: string, style: RichTextStyle): RichTextRun[] | null {
    if (!text || !hasTextStyle(style)) {
      return null;
    }
    return [{
      text,
      ...style
    }];
  }

  function parseInlineRichTextRuns(
    cellElement: Element,
    cellStyle: RichTextStyle,
    deps: Pick<WorksheetParserDependencies, "getTextContent">
  ): RichTextRun[] | null {
    const inlineStringElement = cellElement.getElementsByTagName("is")[0] || null;
    if (!inlineStringElement) {
      return null;
    }
    const runElements = Array.from(inlineStringElement.childNodes).filter((node): node is Element => (
      node.nodeType === Node.ELEMENT_NODE && (node as Element).localName === "r"
    ));
    if (runElements.length === 0) {
      return null;
    }
    return mergeAdjacentRuns(runElements.map((runElement) => ({
      text: Array.from(runElement.getElementsByTagName("t")).map((node) => deps.getTextContent(node)).join(""),
      ...mergeTextStyle(cellStyle, parseRichTextStyle(runElement.getElementsByTagName("rPr")[0] || null))
    })));
  }

  function extractCellOutputValue(
    cellElement: Element,
    sharedStrings: SharedStringEntry[],
    cellStyle: CellStyleInfo,
    deps: Pick<WorksheetParserDependencies, "getTextContent" | "formatCellDisplayValue">,
    formulaOverride = ""
  ): ExtractedCellOutput {
    const type = (cellElement.getAttribute("t") || "").trim();
    const valueNode = cellElement.getElementsByTagName("v")[0] || null;
    const valueText = deps.getTextContent(valueNode);
    const formulaText = formulaOverride || deps.getTextContent(cellElement.getElementsByTagName("f")[0]);
    const cachedValueState: CachedValueState = !formulaText
      ? null
      : !valueNode
        ? "absent"
        : valueText === ""
          ? "present_empty"
          : "present_nonempty";
    if (formulaText) {
      const normalizedFormula = formulaText.startsWith("=") ? formulaText : `=${formulaText}`;
      if (/\[[^\]]+\.xlsx\]/i.test(normalizedFormula)) {
        return {
          valueType: type || "formula",
          rawValue: valueText || normalizedFormula,
          outputValue: normalizedFormula,
          formulaText: normalizedFormula,
          resolutionStatus: "unsupported_external",
          resolutionSource: "external_unsupported",
          cachedValueState,
          richTextRuns: null
        };
      }
      if (valueNode) {
        const formattedValue = deps.formatCellDisplayValue(valueText, cellStyle);
        return {
          valueType: type || "formula",
          rawValue: valueText,
          outputValue: formattedValue ?? valueText,
          formulaText: normalizedFormula,
          resolutionStatus: "resolved",
          resolutionSource: "cached_value",
          cachedValueState,
          richTextRuns: null
        };
      }
      return {
        valueType: type || "formula",
        rawValue: normalizedFormula,
        outputValue: normalizedFormula,
        formulaText: normalizedFormula,
        resolutionStatus: "fallback_formula",
        resolutionSource: "formula_text",
        cachedValueState,
        richTextRuns: null
      };
    }

    if (type === "s") {
      const sharedIndex = Number(valueText || 0);
      const sharedEntry = sharedStrings[sharedIndex] || { text: "", runs: null };
      return {
        valueType: type,
        rawValue: valueText,
        outputValue: sharedEntry.text,
        formulaText: "",
        resolutionStatus: null,
        resolutionSource: null,
        cachedValueState: null,
        richTextRuns: sharedEntry.runs
          ? mergeAdjacentRuns(sharedEntry.runs.map((run) => ({
            text: run.text,
            ...mergeTextStyle(cellStyle.textStyle, run)
          })))
          : createStyledRuns(sharedEntry.text, cellStyle.textStyle)
      };
    }
    if (type === "inlineStr") {
      const inlineText = Array.from(cellElement.getElementsByTagName("t")).map((node) => deps.getTextContent(node)).join("");
      return {
        valueType: type,
        rawValue: inlineText,
        outputValue: inlineText,
        formulaText: "",
        resolutionStatus: null,
        resolutionSource: null,
        cachedValueState: null,
        richTextRuns: parseInlineRichTextRuns(cellElement, cellStyle.textStyle, deps) || createStyledRuns(inlineText, cellStyle.textStyle)
      };
    }
    if (type === "b") {
      return {
        valueType: type,
        rawValue: valueText,
        outputValue: valueText === "1" ? "TRUE" : "FALSE",
        formulaText: "",
        resolutionStatus: null,
        resolutionSource: null,
        cachedValueState: null,
        richTextRuns: createStyledRuns(valueText === "1" ? "TRUE" : "FALSE", cellStyle.textStyle)
      };
    }
    if (type === "str" || type === "e") {
      return {
        valueType: type,
        rawValue: valueText,
        outputValue: valueText,
        formulaText: "",
        resolutionStatus: null,
        resolutionSource: null,
        cachedValueState: null,
        richTextRuns: createStyledRuns(valueText, cellStyle.textStyle)
      };
    }
    if (valueText) {
      const formattedValue = deps.formatCellDisplayValue(valueText, cellStyle);
      if (formattedValue !== null) {
        return {
          valueType: type,
          rawValue: valueText,
          outputValue: formattedValue,
          formulaText: "",
          resolutionStatus: null,
          resolutionSource: null,
          cachedValueState: null,
          richTextRuns: createStyledRuns(formattedValue, cellStyle.textStyle)
        };
      }
    }
    return {
      valueType: type,
      rawValue: valueText,
      outputValue: valueText,
      formulaText: "",
      resolutionStatus: null,
      resolutionSource: null,
      cachedValueState: null,
      richTextRuns: createStyledRuns(valueText, cellStyle.textStyle)
    };
  }

  function shiftReferenceAddress(addressText: string, rowOffset: number, colOffset: number, deps: { lettersToCol: (letters: string) => number; colToLetters: (col: number) => string }): string {
    const match = String(addressText || "").match(/^(\$?)([A-Z]+)(\$?)(\d+)$/i);
    if (!match) return addressText;
    const colAbsolute = match[1] === "$";
    const rowAbsolute = match[3] === "$";
    const baseCol = deps.lettersToCol(match[2]);
    const baseRow = Number(match[4]);
    const shiftedCol = colAbsolute ? baseCol : baseCol + colOffset;
    const shiftedRow = rowAbsolute ? baseRow : baseRow + rowOffset;
    const safeCol = Math.max(1, shiftedCol);
    const safeRow = Math.max(1, shiftedRow);
    return `${colAbsolute ? "$" : ""}${deps.colToLetters(safeCol)}${rowAbsolute ? "$" : ""}${safeRow}`;
  }

  function translateSharedFormula(
    baseFormulaText: string,
    baseAddress: string,
    targetAddress: string,
    deps: {
      parseCellAddress: (address: string) => { row: number; col: number };
      lettersToCol: (letters: string) => number;
      colToLetters: (col: number) => string;
    }
  ): string {
    const basePos = deps.parseCellAddress(baseAddress);
    const targetPos = deps.parseCellAddress(targetAddress);
    if (!basePos.row || !basePos.col || !targetPos.row || !targetPos.col) {
      return baseFormulaText;
    }
    const rowOffset = targetPos.row - basePos.row;
    const colOffset = targetPos.col - basePos.col;
    const normalized = String(baseFormulaText || "").replace(/^=/, "");
    const translated = normalized.replace(
      /(?:'((?:[^']|'')+)'|([A-Za-z0-9_ ]+))!(\$?[A-Z]+\$?\d+)|(\$?[A-Z]+\$?\d+)/g,
      (full, quotedSheet, plainSheet, qualifiedAddress, localAddress) => {
        const address = qualifiedAddress || localAddress;
        if (!address) return full;
        const shifted = shiftReferenceAddress(address, rowOffset, colOffset, deps);
        if (qualifiedAddress) {
          const sheetPrefix = quotedSheet ? `'${quotedSheet}'` : plainSheet;
          return `${sheetPrefix}!${shifted}`;
        }
        return shifted;
      }
    );
    return translated.startsWith("=") ? translated : `=${translated}`;
  }

  function parseWorksheet(
    files: Map<string, Uint8Array>,
    sheetName: string,
    sheetPath: string,
    sheetIndex: number,
    sharedStrings: SharedStringEntry[],
    cellStyles: CellStyleInfo[],
    deps: WorksheetParserDependencies & { lettersToCol: (letters: string) => number; colToLetters: (col: number) => string }
  ): ParsedSheet {
    const bytes = files.get(sheetPath);
    if (!bytes) {
      throw new Error(`Sheet XML not found: ${sheetPath}`);
    }
    const doc = deps.xmlToDocument(deps.decodeXmlText(bytes));
    const sharedFormulaMap = new Map<string, { address: string; formulaText: string }>();
    const hyperlinks = parseWorksheetHyperlinks(files, doc, sheetPath, deps);
    const cells = Array.from(doc.getElementsByTagName("c")).map((cellElement) => {
      const address = cellElement.getAttribute("r") || "";
      const position = deps.parseCellAddress(address);
      const styleIndex = Number(cellElement.getAttribute("s") || 0);
      const cellStyle = cellStyles[styleIndex] || {
        borders: deps.EMPTY_BORDERS,
        numFmtId: 0,
        formatCode: "General",
        textStyle: {
          bold: false,
          italic: false,
          strike: false,
          underline: false
        }
      };
      let formulaOverride = "";
      const formulaElement = cellElement.getElementsByTagName("f")[0] || null;
      const formulaType = formulaElement?.getAttribute("t") || "";
      const spillRef = formulaElement?.getAttribute("ref") || "";
      const sharedIndex = formulaElement?.getAttribute("si") || "";
      const formulaText = deps.getTextContent(formulaElement);
      if (formulaType === "shared" && sharedIndex) {
        if (formulaText) {
          const normalizedFormula = formulaText.startsWith("=") ? formulaText : `=${formulaText}`;
          sharedFormulaMap.set(sharedIndex, { address, formulaText: normalizedFormula });
          formulaOverride = normalizedFormula;
        } else {
          const sharedBase = sharedFormulaMap.get(sharedIndex);
          if (sharedBase) {
            formulaOverride = translateSharedFormula(sharedBase.formulaText, sharedBase.address, address, deps);
          }
        }
      }
      const output = extractCellOutputValue(cellElement, sharedStrings, cellStyle, deps, formulaOverride);
      return {
        address,
        row: position.row,
        col: position.col,
        valueType: output.valueType,
        rawValue: output.rawValue,
        outputValue: output.outputValue,
        formulaText: output.formulaText,
        resolutionStatus: output.resolutionStatus,
        resolutionSource: output.resolutionSource,
        cachedValueState: output.cachedValueState,
        styleIndex,
        borders: cellStyle.borders,
        numFmtId: cellStyle.numFmtId,
        formatCode: cellStyle.formatCode,
        textStyle: cellStyle.textStyle,
        richTextRuns: output.richTextRuns,
        formulaType,
        spillRef,
        hyperlink: hyperlinks.get(address) || null
      } satisfies ParsedCell;
    });
    const merges = Array.from(doc.getElementsByTagName("mergeCell")).map((mergeElement) => deps.parseRangeRef(mergeElement.getAttribute("ref") || ""));
    const tables = deps.parseWorksheetTables(files, doc, sheetName, sheetPath);
    const assetDeps = deps.buildAssetDeps();
    const images = deps.parseDrawingImages(files, sheetName, sheetPath, assetDeps);
    const charts = deps.parseDrawingCharts(files, sheetName, sheetPath, assetDeps);
    const shapes = deps.parseDrawingShapes(files, sheetName, sheetPath, assetDeps);
    let maxRow = 0;
    let maxCol = 0;
    for (const cell of cells) {
      if (cell.row > maxRow) maxRow = cell.row;
      if (cell.col > maxCol) maxCol = cell.col;
    }
    for (const merge of merges) {
      if (merge.endRow > maxRow) maxRow = merge.endRow;
      if (merge.endCol > maxCol) maxCol = merge.endCol;
    }
    return {
      name: sheetName,
      index: sheetIndex,
      path: sheetPath,
      cells,
      merges,
      tables,
      images,
      charts,
      shapes,
      maxRow,
      maxCol
    };
  }

  const worksheetParserApi = {
    extractCellOutputValue,
    expandRangeAddresses,
    parseWorksheetHyperlinks,
    shiftReferenceAddress,
    translateSharedFormula,
    parseWorksheet
  };

  moduleRegistry.registerModule("worksheetParser", worksheetParserApi);
})();
