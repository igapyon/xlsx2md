(() => {
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

  type ShapeBlock = {
    startRow: number;
    startCol: number;
    endRow: number;
    endCol: number;
    shapeIndexes: number[];
  };

  type DrawingHelper = {
    renderShapeSvg?: (shapeNode: Element, anchor: Element, sheetName: string, shapeIndex: number) => {
      filename: string;
      path: string;
      data: Uint8Array;
    } | null;
  } | null;

  type ParseAssetDependencies = {
    parseRelationships: (files: Map<string, Uint8Array>, relsPath: string, sourcePath: string) => Map<string, string>;
    buildRelsPath: (sourcePath: string) => string;
    xmlToDocument: (xmlText: string) => Document;
    decodeXmlText: (bytes: Uint8Array) => string;
    getElementsByLocalName: (root: ParentNode, localName: string) => Element[];
    getFirstChildByLocalName: (root: ParentNode, localName: string) => Element | null;
    getDirectChildByLocalName: (root: ParentNode | null, localName: string) => Element | null;
    getTextContent: (node: Element | null | undefined) => string;
    colToLetters: (col: number) => string;
    drawingHelper: DrawingHelper;
    defaultCellWidthEmu: number;
    defaultCellHeightEmu: number;
    shapeBlockGapXEmu: number;
    shapeBlockGapYEmu: number;
  };

  function createSafeSheetAssetDir(sheetName: string): string {
    return sheetName.replace(/[\\/:*?"<>|]+/g, "_").trim() || "Sheet";
  }

  function getImageExtension(mediaPath: string): string {
    const match = mediaPath.match(/\.([a-z0-9]+)$/i);
    return match ? match[1].toLowerCase() : "bin";
  }

  function parseDrawingImages(
    files: Map<string, Uint8Array>,
    sheetName: string,
    sheetPath: string,
    deps: ParseAssetDependencies
  ): ParsedImageAsset[] {
    const sheetRels = deps.parseRelationships(files, deps.buildRelsPath(sheetPath), sheetPath);
    const imageAssets: ParsedImageAsset[] = [];
    let imageCounter = 1;

    for (const drawingPath of sheetRels.values()) {
      if (!/\/drawings\/.+\.xml$/i.test(drawingPath)) continue;
      const drawingBytes = files.get(drawingPath);
      if (!drawingBytes) continue;
      const drawingDoc = deps.xmlToDocument(deps.decodeXmlText(drawingBytes));
      const drawingRels = deps.parseRelationships(files, deps.buildRelsPath(drawingPath), drawingPath);
      const anchors = deps.getElementsByLocalName(drawingDoc, "oneCellAnchor").concat(deps.getElementsByLocalName(drawingDoc, "twoCellAnchor"));

      for (const anchor of anchors) {
        const from = deps.getFirstChildByLocalName(anchor, "from");
        const colNode = deps.getFirstChildByLocalName(from || anchor, "col");
        const rowNode = deps.getFirstChildByLocalName(from || anchor, "row");
        const col = Number(deps.getTextContent(colNode)) + 1;
        const row = Number(deps.getTextContent(rowNode)) + 1;
        if (!Number.isFinite(col) || !Number.isFinite(row) || col <= 0 || row <= 0) {
          continue;
        }

        const blip = deps.getElementsByLocalName(anchor, "blip")[0] || null;
        const embedId = blip?.getAttribute("r:embed") || blip?.getAttribute("embed") || "";
        const mediaPath = drawingRels.get(embedId) || "";
        if (!mediaPath) continue;
        const mediaBytes = files.get(mediaPath);
        if (!mediaBytes) continue;

        const extension = getImageExtension(mediaPath);
        const safeDir = createSafeSheetAssetDir(sheetName);
        const filename = `image_${String(imageCounter).padStart(3, "0")}.${extension}`;
        imageAssets.push({
          sheetName,
          filename,
          path: `assets/${safeDir}/${filename}`,
          anchor: `${deps.colToLetters(col)}${row}`,
          data: new Uint8Array(mediaBytes),
          mediaPath
        });
        imageCounter += 1;
      }
    }

    return imageAssets;
  }

  function parseChartType(chartDoc: Document, deps: ParseAssetDependencies): string {
    const typeMap: Array<{ localName: string; label: string }> = [
      { localName: "barChart", label: "Bar Chart" },
      { localName: "lineChart", label: "Line Chart" },
      { localName: "pieChart", label: "Pie Chart" },
      { localName: "doughnutChart", label: "Doughnut Chart" },
      { localName: "areaChart", label: "Area Chart" },
      { localName: "scatterChart", label: "Scatter Chart" },
      { localName: "radarChart", label: "Radar Chart" },
      { localName: "bubbleChart", label: "Bubble Chart" }
    ];
    const matched = typeMap
      .filter((entry) => deps.getElementsByLocalName(chartDoc, entry.localName).length > 0)
      .map((entry) => entry.label);
    if (matched.length === 0) return "Chart";
    if (matched.length === 1) return matched[0];
    return `${matched.join(" + ")} (Combined)`;
  }

  function parseChartTitle(chartDoc: Document, deps: ParseAssetDependencies): string {
    const richText = deps.getElementsByLocalName(chartDoc, "t")
      .map((node) => deps.getTextContent(node))
      .filter(Boolean);
    if (richText.length > 0) {
      return richText.join("").trim();
    }
    return "";
  }

  function parseChartSeries(chartDoc: Document, deps: ParseAssetDependencies): ParsedChartAsset["series"] {
    const plotArea = deps.getFirstChildByLocalName(chartDoc, "plotArea") || chartDoc.documentElement;
    const axisPositionById = new Map<string, string>();
    for (const axisNode of deps.getElementsByLocalName(plotArea, "valAx")) {
      const axisIdNode = deps.getFirstChildByLocalName(axisNode, "axId");
      const axisPosNode = deps.getFirstChildByLocalName(axisNode, "axPos");
      const axisId = axisIdNode?.getAttribute("val") || deps.getTextContent(axisIdNode);
      const axisPos = axisPosNode?.getAttribute("val") || deps.getTextContent(axisPosNode);
      if (axisId) {
        axisPositionById.set(axisId, axisPos || "");
      }
    }

    const chartContainerNames = [
      "barChart",
      "lineChart",
      "pieChart",
      "doughnutChart",
      "areaChart",
      "scatterChart",
      "radarChart",
      "bubbleChart"
    ];
    const series: ParsedChartAsset["series"] = [];

    for (const localName of chartContainerNames) {
      for (const chartNode of deps.getElementsByLocalName(plotArea, localName)) {
        const axisIds = deps.getElementsByLocalName(chartNode, "axId")
          .map((node) => node.getAttribute("val") || deps.getTextContent(node))
          .filter(Boolean);
        const isSecondary = axisIds.some((axisId) => axisPositionById.get(axisId) === "r");

        for (const seriesNode of deps.getElementsByLocalName(chartNode, "ser")) {
          const txNode = deps.getFirstChildByLocalName(seriesNode, "tx") || seriesNode;
          const nameRef = deps.getFirstChildByLocalName(txNode, "f");
          const nameValue = deps.getFirstChildByLocalName(txNode, "v");
          const nameText = deps.getElementsByLocalName(txNode, "t")
            .map((node) => deps.getTextContent(node))
            .join("")
            .trim();
          const catRef = deps.getFirstChildByLocalName(deps.getFirstChildByLocalName(deps.getFirstChildByLocalName(seriesNode, "cat") || seriesNode, "strRef") || seriesNode, "f")
            || deps.getFirstChildByLocalName(deps.getFirstChildByLocalName(deps.getFirstChildByLocalName(seriesNode, "cat") || seriesNode, "numRef") || seriesNode, "f");
          const valRef = deps.getFirstChildByLocalName(deps.getFirstChildByLocalName(seriesNode, "val") || seriesNode, "f")
            || deps.getFirstChildByLocalName(deps.getFirstChildByLocalName(deps.getFirstChildByLocalName(seriesNode, "val") || seriesNode, "numRef") || seriesNode, "f");
          series.push({
            name: nameText || deps.getTextContent(nameValue) || deps.getTextContent(nameRef) || "Series",
            categoriesRef: deps.getTextContent(catRef),
            valuesRef: deps.getTextContent(valRef),
            axis: isSecondary ? "secondary" : "primary"
          });
        }
      }
    }

    return series;
  }

  function parseDrawingCharts(
    files: Map<string, Uint8Array>,
    sheetName: string,
    sheetPath: string,
    deps: ParseAssetDependencies
  ): ParsedChartAsset[] {
    const sheetRels = deps.parseRelationships(files, deps.buildRelsPath(sheetPath), sheetPath);
    const charts: ParsedChartAsset[] = [];

    for (const drawingPath of sheetRels.values()) {
      if (!/\/drawings\/.+\.xml$/i.test(drawingPath)) continue;
      const drawingBytes = files.get(drawingPath);
      if (!drawingBytes) continue;
      const drawingDoc = deps.xmlToDocument(deps.decodeXmlText(drawingBytes));
      const drawingRels = deps.parseRelationships(files, deps.buildRelsPath(drawingPath), drawingPath);
      const anchors = deps.getElementsByLocalName(drawingDoc, "oneCellAnchor").concat(deps.getElementsByLocalName(drawingDoc, "twoCellAnchor"));

      for (const anchor of anchors) {
        const from = deps.getFirstChildByLocalName(anchor, "from");
        const colNode = deps.getFirstChildByLocalName(from || anchor, "col");
        const rowNode = deps.getFirstChildByLocalName(from || anchor, "row");
        const col = Number(deps.getTextContent(colNode)) + 1;
        const row = Number(deps.getTextContent(rowNode)) + 1;
        if (!Number.isFinite(col) || !Number.isFinite(row) || col <= 0 || row <= 0) {
          continue;
        }

        const chartNode = deps.getFirstChildByLocalName(anchor, "graphicFrame");
        const chartRef = deps.getElementsByLocalName(chartNode || anchor, "chart")[0] || null;
        const relId = chartRef?.getAttribute("r:id") || chartRef?.getAttribute("id") || "";
        const chartPath = drawingRels.get(relId) || "";
        if (!chartPath) continue;
        const chartBytes = files.get(chartPath);
        if (!chartBytes) continue;
        const chartDoc = deps.xmlToDocument(deps.decodeXmlText(chartBytes));

        charts.push({
          sheetName,
          anchor: `${deps.colToLetters(col)}${row}`,
          chartPath,
          title: parseChartTitle(chartDoc, deps),
          chartType: parseChartType(chartDoc, deps),
          series: parseChartSeries(chartDoc, deps)
        });
      }
    }

    return charts;
  }

  function parseShapeKind(shapeNode: Element | null, deps: ParseAssetDependencies): string {
    if (!shapeNode) return "Shape";
    if (shapeNode.localName === "cxnSp") {
      const geomNode = deps.getFirstChildByLocalName(deps.getFirstChildByLocalName(shapeNode, "spPr") || shapeNode, "prstGeom");
      const prst = String(geomNode?.getAttribute("prst") || "").trim();
      if (prst === "straightConnector1") {
        return "Straight Arrow Connector";
      }
      return prst ? `Connector (${prst})` : "Connector";
    }
    if (shapeNode.localName !== "sp") {
      return "Shape";
    }
    const nvSpPr = deps.getFirstChildByLocalName(shapeNode, "nvSpPr");
    const cNvSpPr = deps.getFirstChildByLocalName(nvSpPr || shapeNode, "cNvSpPr");
    if (cNvSpPr?.getAttribute("txBox") === "1") {
      return "Text Box";
    }
    const geomNode = deps.getFirstChildByLocalName(deps.getFirstChildByLocalName(shapeNode, "spPr") || shapeNode, "prstGeom");
    const prst = String(geomNode?.getAttribute("prst") || "").trim();
    if (prst === "rect") {
      return "Rectangle";
    }
    return prst ? `Shape (${prst})` : "Shape";
  }

  function parseShapeText(shapeNode: Element | null, deps: ParseAssetDependencies): string {
    return deps.getElementsByLocalName(shapeNode || document, "t")
      .map((node) => deps.getTextContent(node))
      .filter(Boolean)
      .join("")
      .trim();
  }

  function parseShapeExt(anchor: Element, shapeNode: Element | null, deps: ParseAssetDependencies): { widthEmu: number | null; heightEmu: number | null } {
    const extNode = deps.getDirectChildByLocalName(anchor, "ext")
      || deps.getDirectChildByLocalName(
        deps.getDirectChildByLocalName(deps.getDirectChildByLocalName(shapeNode || anchor, "spPr") || shapeNode || anchor, "xfrm"),
        "ext"
      );
    const widthEmu = Number(extNode?.getAttribute("cx") || "");
    const heightEmu = Number(extNode?.getAttribute("cy") || "");
    return {
      widthEmu: Number.isFinite(widthEmu) ? widthEmu : null,
      heightEmu: Number.isFinite(heightEmu) ? heightEmu : null
    };
  }

  function flattenXmlNodeEntries(
    node: Element | null,
    deps: ParseAssetDependencies,
    path = "",
    entries: { key: string; value: string }[] = []
  ): { key: string; value: string }[] {
    if (!node) return entries;
    const nodeName = node.tagName || node.nodeName || node.localName || "node";
    const currentPath = path ? `${path}/${nodeName}` : nodeName;

    for (const attribute of Array.from(node.attributes)) {
      entries.push({
        key: `${currentPath}@${attribute.name}`,
        value: attribute.value
      });
    }

    const directText = Array.from(node.childNodes)
      .filter((child) => child.nodeType === Node.TEXT_NODE)
      .map((child) => (child.textContent || "").trim())
      .filter(Boolean)
      .join(" ");
    if (directText) {
      entries.push({
        key: `${currentPath}#text`,
        value: directText
      });
    }

    for (const child of Array.from(node.childNodes)) {
      if (child.nodeType === Node.ELEMENT_NODE) {
        flattenXmlNodeEntries(child as Element, deps, currentPath, entries);
      }
    }
    return entries;
  }

  function parseShapeRawEntries(anchor: Element, deps: ParseAssetDependencies): { key: string; value: string }[] {
    const entries: { key: string; value: string }[] = [];
    return flattenXmlNodeEntries(anchor, deps, "", entries);
  }

  function renderHierarchicalRawEntries(entries: { key: string; value: string }[]): string[] {
    type RawTreeNode = {
      children: Map<string, RawTreeNode>;
      value: string | null;
    };

    const root: RawTreeNode = {
      children: new Map<string, RawTreeNode>(),
      value: null
    };

    for (const entry of entries) {
      const parts = entry.key.split("/").filter(Boolean);
      let current = root;
      for (const part of parts) {
        if (!current.children.has(part)) {
          current.children.set(part, {
            children: new Map<string, RawTreeNode>(),
            value: null
          });
        }
        current = current.children.get(part)!;
      }
      current.value = entry.value;
    }

    const lines: string[] = [];

    function visit(node: RawTreeNode, depth: number): void {
      for (const [key, child] of node.children.entries()) {
        const indent = " ".repeat(depth * 4);
        if (child.value !== null) {
          lines.push(`${indent}- \`${key}\`: \`${child.value}\``);
        } else {
          lines.push(`${indent}- \`${key}\``);
        }
        visit(child, depth + 1);
      }
    }

    visit(root, 0);
    return lines;
  }

  function parseAnchorInt(anchor: Element | null, parentName: string, childName: string, deps: ParseAssetDependencies): number | null {
    const parent = deps.getFirstChildByLocalName(anchor || document, parentName);
    const child = deps.getFirstChildByLocalName(parent || anchor || document, childName);
    const value = Number(deps.getTextContent(child));
    return Number.isFinite(value) ? value : null;
  }

  function parseShapeBoundingBox(anchor: Element, shapeNode: Element | null, widthEmu: number | null, heightEmu: number | null, deps: ParseAssetDependencies): {
    left: number;
    top: number;
    right: number;
    bottom: number;
  } {
    const fromCol = parseAnchorInt(anchor, "from", "col", deps) || 0;
    const fromRow = parseAnchorInt(anchor, "from", "row", deps) || 0;
    const fromColOff = parseAnchorInt(anchor, "from", "colOff", deps) || 0;
    const fromRowOff = parseAnchorInt(anchor, "from", "rowOff", deps) || 0;
    const toCol = parseAnchorInt(anchor, "to", "col", deps);
    const toRow = parseAnchorInt(anchor, "to", "row", deps);
    const toColOff = parseAnchorInt(anchor, "to", "colOff", deps) || 0;
    const toRowOff = parseAnchorInt(anchor, "to", "rowOff", deps) || 0;

    const left = fromCol * deps.defaultCellWidthEmu + fromColOff;
    const top = fromRow * deps.defaultCellHeightEmu + fromRowOff;

    if (toCol !== null && toRow !== null) {
      return {
        left,
        top,
        right: toCol * deps.defaultCellWidthEmu + toColOff,
        bottom: toRow * deps.defaultCellHeightEmu + toRowOff
      };
    }

    const ext = parseShapeExt(anchor, shapeNode, deps);
    return {
      left,
      top,
      right: left + Math.max(1, ext.widthEmu || widthEmu || deps.defaultCellWidthEmu),
      bottom: top + Math.max(1, ext.heightEmu || heightEmu || deps.defaultCellHeightEmu)
    };
  }

  function bboxGap(a: ParsedShapeAsset["bbox"], b: ParsedShapeAsset["bbox"]): { dx: number; dy: number } {
    const dx = a.right < b.left
      ? b.left - a.right
      : b.right < a.left
        ? a.left - b.right
        : 0;
    const dy = a.bottom < b.top
      ? b.top - a.bottom
      : b.bottom < a.top
        ? a.top - b.bottom
        : 0;
    return { dx, dy };
  }

  function extractShapeBlocks(shapes: ParsedShapeAsset[], deps: Pick<ParseAssetDependencies, "defaultCellWidthEmu" | "defaultCellHeightEmu" | "shapeBlockGapXEmu" | "shapeBlockGapYEmu">): ShapeBlock[] {
    if (shapes.length === 0) return [];
    const visited = new Array(shapes.length).fill(false);
    const blocks: ShapeBlock[] = [];

    for (let i = 0; i < shapes.length; i += 1) {
      if (visited[i]) continue;
      const queue = [i];
      visited[i] = true;
      const shapeIndexes: number[] = [];

      while (queue.length > 0) {
        const currentIndex = queue.shift()!;
        shapeIndexes.push(currentIndex);
        const current = shapes[currentIndex];
        for (let j = 0; j < shapes.length; j += 1) {
          if (visited[j]) continue;
          const other = shapes[j];
          const { dx, dy } = bboxGap(current.bbox, other.bbox);
          if (dx <= deps.shapeBlockGapXEmu && dy <= deps.shapeBlockGapYEmu) {
            visited[j] = true;
            queue.push(j);
          }
        }
      }

      let minLeft = Number.POSITIVE_INFINITY;
      let minTop = Number.POSITIVE_INFINITY;
      let maxRight = 0;
      let maxBottom = 0;
      for (const index of shapeIndexes) {
        const bbox = shapes[index].bbox;
        minLeft = Math.min(minLeft, bbox.left);
        minTop = Math.min(minTop, bbox.top);
        maxRight = Math.max(maxRight, bbox.right);
        maxBottom = Math.max(maxBottom, bbox.bottom);
      }
      blocks.push({
        startCol: Math.floor(minLeft / deps.defaultCellWidthEmu) + 1,
        startRow: Math.floor(minTop / deps.defaultCellHeightEmu) + 1,
        endCol: Math.floor(maxRight / deps.defaultCellWidthEmu) + 1,
        endRow: Math.floor(maxBottom / deps.defaultCellHeightEmu) + 1,
        shapeIndexes: shapeIndexes.sort((a, b) => a - b)
      });
    }

    return blocks.sort((a, b) => (a.startRow - b.startRow) || (a.startCol - b.startCol));
  }

  function parseDrawingShapes(
    files: Map<string, Uint8Array>,
    sheetName: string,
    sheetPath: string,
    deps: ParseAssetDependencies
  ): ParsedShapeAsset[] {
    const sheetRels = deps.parseRelationships(files, deps.buildRelsPath(sheetPath), sheetPath);
    const shapes: ParsedShapeAsset[] = [];
    let shapeCounter = 1;

    for (const drawingPath of sheetRels.values()) {
      if (!/\/drawings\/.+\.xml$/i.test(drawingPath)) continue;
      const drawingBytes = files.get(drawingPath);
      if (!drawingBytes) continue;
      const drawingDoc = deps.xmlToDocument(deps.decodeXmlText(drawingBytes));
      const anchors = deps.getElementsByLocalName(drawingDoc, "oneCellAnchor").concat(deps.getElementsByLocalName(drawingDoc, "twoCellAnchor"));

      for (const anchor of anchors) {
        const from = deps.getFirstChildByLocalName(anchor, "from");
        const colNode = deps.getFirstChildByLocalName(from || anchor, "col");
        const rowNode = deps.getFirstChildByLocalName(from || anchor, "row");
        const col = Number(deps.getTextContent(colNode)) + 1;
        const row = Number(deps.getTextContent(rowNode)) + 1;
        if (!Number.isFinite(col) || !Number.isFinite(row) || col <= 0 || row <= 0) {
          continue;
        }

        if (deps.getElementsByLocalName(anchor, "blip").length > 0) continue;
        if (deps.getElementsByLocalName(anchor, "chart").length > 0) continue;

        const shapeNode = deps.getFirstChildByLocalName(anchor, "sp") || deps.getFirstChildByLocalName(anchor, "cxnSp");
        if (!shapeNode) continue;
        const cNvPr = deps.getFirstChildByLocalName(deps.getFirstChildByLocalName(shapeNode, shapeNode.localName === "sp" ? "nvSpPr" : "nvCxnSpPr") || shapeNode, "cNvPr");
        const { widthEmu, heightEmu } = parseShapeExt(anchor, shapeNode, deps);
        const svgAsset = deps.drawingHelper?.renderShapeSvg?.(shapeNode, anchor, sheetName, shapeCounter) || null;
        shapes.push({
          sheetName,
          anchor: `${deps.colToLetters(col)}${row}`,
          name: String(cNvPr?.getAttribute("name") || "").trim() || "Shape",
          kind: parseShapeKind(shapeNode, deps),
          text: parseShapeText(shapeNode, deps),
          widthEmu,
          heightEmu,
          elementName: `xdr:${shapeNode.localName}`,
          anchorElementName: anchor.tagName || anchor.nodeName || anchor.localName || "anchor",
          rawEntries: parseShapeRawEntries(anchor, deps),
          bbox: parseShapeBoundingBox(anchor, shapeNode, widthEmu, heightEmu, deps),
          svgFilename: svgAsset?.filename || null,
          svgPath: svgAsset?.path || null,
          svgData: svgAsset?.data || null
        });
        shapeCounter += 1;
      }
    }

    return shapes;
  }

  (globalThis as typeof globalThis & {
    __xlsx2mdSheetAssets?: {
      createSafeSheetAssetDir: typeof createSafeSheetAssetDir;
      parseDrawingImages: typeof parseDrawingImages;
      parseDrawingCharts: typeof parseDrawingCharts;
      parseDrawingShapes: typeof parseDrawingShapes;
      extractShapeBlocks: typeof extractShapeBlocks;
      renderHierarchicalRawEntries: typeof renderHierarchicalRawEntries;
    };
  }).__xlsx2mdSheetAssets = {
    createSafeSheetAssetDir,
    parseDrawingImages,
    parseDrawingCharts,
    parseDrawingShapes,
    extractShapeBlocks,
    renderHierarchicalRawEntries
  };
})();
