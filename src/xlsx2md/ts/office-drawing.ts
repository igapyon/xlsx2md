/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */

(() => {
  const moduleRegistry = getXlsx2mdModuleRegistry();
  type SvgRenderResult = {
    filename: string;
    path: string;
    data: Uint8Array;
  } | null;

  const textEncoder = new TextEncoder();
  const runtimeEnv = requireXlsx2mdRuntimeEnv();

  function getDirectChildByLocalName(root: ParentNode | null, localName: string): Element | null {
    if (!root) return null;
    for (const node of Array.from(root.childNodes)) {
      if (node.nodeType === runtimeEnv.ELEMENT_NODE && (node as Element).localName === localName) {
        return node as Element;
      }
    }
    return null;
  }

  function getElementsByLocalName(root: ParentNode | null, localName: string): Element[] {
    if (!root) return [];
    const elements = Array.from(root.getElementsByTagName("*"));
    return elements.filter((element) => element.localName === localName);
  }

  function getTextContent(node: Element | null | undefined): string {
    return (node?.textContent || "").replace(/\r\n/g, "\n");
  }

  function createSafeSheetAssetDir(sheetName: string): string {
    return sheetName.replace(/[\\/:*?"<>|]+/g, "_").trim() || "Sheet";
  }

  function escapeXml(text: string): string {
    return String(text || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");
  }

  function emuToPx(emu: number | null, fallback: number): number {
    if (!Number.isFinite(emu as number) || (emu as number) <= 0) return fallback;
    return Math.max(1, Math.round((emu as number) / 9525));
  }

  function parseHexColor(root: Element | null): string | null {
    const srgb = getElementsByLocalName(root, "srgbClr")[0] || null;
    if (srgb?.getAttribute("val")) {
      return `#${String(srgb.getAttribute("val")).trim()}`;
    }
    const scheme = getElementsByLocalName(root, "schemeClr")[0] || null;
    const schemeVal = String(scheme?.getAttribute("val") || "").trim();
    const schemeMap: Record<string, string> = {
      accent1: "#4472C4",
      accent2: "#ED7D31",
      accent3: "#A5A5A5",
      accent4: "#FFC000",
      accent5: "#5B9BD5",
      accent6: "#70AD47",
      tx1: "#000000",
      tx2: "#44546A",
      lt1: "#FFFFFF",
      lt2: "#E7E6E6"
    };
    return schemeMap[schemeVal] || null;
  }

  function parseShapeText(shapeNode: Element | null): string {
    return getElementsByLocalName(shapeNode, "t")
      .map((node) => getTextContent(node).trim())
      .filter(Boolean)
      .join("\n")
      .trim();
  }

  function parseShapeKind(shapeNode: Element | null): "textbox" | "rect" | "connector" | null {
    if (!shapeNode) return null;
    if (shapeNode.localName === "cxnSp") {
      return "connector";
    }
    if (shapeNode.localName !== "sp") {
      return null;
    }
    const nvSpPr = getDirectChildByLocalName(shapeNode, "nvSpPr");
    const cNvSpPr = getDirectChildByLocalName(nvSpPr, "cNvSpPr");
    if (cNvSpPr?.getAttribute("txBox") === "1") {
      return "textbox";
    }
    const spPr = getDirectChildByLocalName(shapeNode, "spPr");
    const prstGeom = getDirectChildByLocalName(spPr, "prstGeom");
    if (String(prstGeom?.getAttribute("prst") || "").trim() === "rect") {
      return "rect";
    }
    return null;
  }

  function parseShapeDimensions(anchor: Element, shapeNode: Element | null): { widthPx: number; heightPx: number } {
    const extNode = getDirectChildByLocalName(anchor, "ext")
      || getDirectChildByLocalName(getDirectChildByLocalName(getDirectChildByLocalName(shapeNode || anchor, "spPr"), "xfrm"), "ext");
    const widthEmu = Number(extNode?.getAttribute("cx") || "");
    const heightEmu = Number(extNode?.getAttribute("cy") || "");
    return {
      widthPx: emuToPx(widthEmu, 160),
      heightPx: emuToPx(heightEmu, 48)
    };
  }

  function renderRectLikeSvg(shapeNode: Element, anchor: Element, text: string, treatAsTextbox: boolean): string {
    const { widthPx, heightPx } = parseShapeDimensions(anchor, shapeNode);
    const spPr = getDirectChildByLocalName(shapeNode, "spPr");
    const fillColor = parseHexColor(getDirectChildByLocalName(spPr, "solidFill")) || (treatAsTextbox ? "#FFFFFF" : "#F3F3F3");
    const lineNode = getDirectChildByLocalName(spPr, "ln");
    const strokeColor = parseHexColor(lineNode) || "#333333";
    const strokeWidth = Math.max(1, Math.round(Number(lineNode?.getAttribute("w") || "") / 9525) || 1);
    const safeText = escapeXml(text);
    const textMarkup = safeText
      ? `<text x="${Math.round(widthPx / 2)}" y="${Math.round(heightPx / 2)}" text-anchor="middle" dominant-baseline="middle" font-size="14" font-family="sans-serif" fill="#000000">${safeText}</text>`
      : "";
    return [
      `<svg xmlns="http://www.w3.org/2000/svg" width="${widthPx}" height="${heightPx}" viewBox="0 0 ${widthPx} ${heightPx}">`,
      `  <rect x="1" y="1" width="${Math.max(1, widthPx - 2)}" height="${Math.max(1, heightPx - 2)}" fill="${fillColor}" stroke="${strokeColor}" stroke-width="${strokeWidth}"/>`,
      textMarkup ? `  ${textMarkup}` : "",
      `</svg>`
    ].filter(Boolean).join("\n");
  }

  function renderConnectorSvg(shapeNode: Element, anchor: Element): string {
    const { widthPx, heightPx } = parseShapeDimensions(anchor, shapeNode);
    const spPr = getDirectChildByLocalName(shapeNode, "spPr");
    const lineNode = getDirectChildByLocalName(spPr, "ln");
    const strokeColor = parseHexColor(lineNode) || "#333333";
    const strokeWidth = Math.max(1, Math.round(Number(lineNode?.getAttribute("w") || "") / 9525) || 1);
    const effectiveHeight = Math.max(heightPx, 24);
    const y = Math.round(effectiveHeight / 2);
    return [
      `<svg xmlns="http://www.w3.org/2000/svg" width="${widthPx}" height="${effectiveHeight}" viewBox="0 0 ${widthPx} ${effectiveHeight}">`,
      `  <defs>`,
      `    <marker id="arrow" markerWidth="10" markerHeight="10" refX="8" refY="3" orient="auto" markerUnits="strokeWidth">`,
      `      <path d="M0,0 L0,6 L9,3 z" fill="${strokeColor}"/>`,
      `    </marker>`,
      `  </defs>`,
      `  <line x1="2" y1="${y}" x2="${Math.max(2, widthPx - 4)}" y2="${y}" stroke="${strokeColor}" stroke-width="${strokeWidth}" marker-end="url(#arrow)"/>`,
      `</svg>`
    ].join("\n");
  }

  function renderShapeSvg(shapeNode: Element, anchor: Element, sheetName: string, shapeIndex: number): SvgRenderResult {
    const kind = parseShapeKind(shapeNode);
    if (!kind) return null;
    let svg = "";
    if (kind === "connector") {
      svg = renderConnectorSvg(shapeNode, anchor);
    } else {
      svg = renderRectLikeSvg(shapeNode, anchor, parseShapeText(shapeNode), kind === "textbox");
    }
    const safeDir = createSafeSheetAssetDir(sheetName);
    const filename = `shape_${String(shapeIndex).padStart(3, "0")}.svg`;
    return {
      filename,
      path: `assets/${safeDir}/${filename}`,
      data: textEncoder.encode(`${svg}\n`)
    };
  }

  const officeDrawingApi = {
    renderShapeSvg
  };

  moduleRegistry.registerModule("officeDrawing", officeDrawingApi);
})();
