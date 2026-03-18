(() => {
    const textEncoder = new TextEncoder();
    function getDirectChildByLocalName(root, localName) {
        if (!root)
            return null;
        for (const node of Array.from(root.childNodes)) {
            if (node.nodeType === Node.ELEMENT_NODE && node.localName === localName) {
                return node;
            }
        }
        return null;
    }
    function getElementsByLocalName(root, localName) {
        if (!root)
            return [];
        const elements = Array.from(root.getElementsByTagName("*"));
        return elements.filter((element) => element.localName === localName);
    }
    function getTextContent(node) {
        return ((node === null || node === void 0 ? void 0 : node.textContent) || "").replace(/\r\n/g, "\n");
    }
    function createSafeSheetAssetDir(sheetName) {
        return sheetName.replace(/[\\/:*?"<>|]+/g, "_").trim() || "Sheet";
    }
    function escapeXml(text) {
        return String(text || "")
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&apos;");
    }
    function emuToPx(emu, fallback) {
        if (!Number.isFinite(emu) || emu <= 0)
            return fallback;
        return Math.max(1, Math.round(emu / 9525));
    }
    function parseHexColor(root) {
        const srgb = getElementsByLocalName(root, "srgbClr")[0] || null;
        if (srgb === null || srgb === void 0 ? void 0 : srgb.getAttribute("val")) {
            return `#${String(srgb.getAttribute("val")).trim()}`;
        }
        const scheme = getElementsByLocalName(root, "schemeClr")[0] || null;
        const schemeVal = String((scheme === null || scheme === void 0 ? void 0 : scheme.getAttribute("val")) || "").trim();
        const schemeMap = {
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
    function parseShapeText(shapeNode) {
        return getElementsByLocalName(shapeNode, "t")
            .map((node) => getTextContent(node).trim())
            .filter(Boolean)
            .join("\n")
            .trim();
    }
    function parseShapeKind(shapeNode) {
        if (!shapeNode)
            return null;
        if (shapeNode.localName === "cxnSp") {
            return "connector";
        }
        if (shapeNode.localName !== "sp") {
            return null;
        }
        const nvSpPr = getDirectChildByLocalName(shapeNode, "nvSpPr");
        const cNvSpPr = getDirectChildByLocalName(nvSpPr, "cNvSpPr");
        if ((cNvSpPr === null || cNvSpPr === void 0 ? void 0 : cNvSpPr.getAttribute("txBox")) === "1") {
            return "textbox";
        }
        const spPr = getDirectChildByLocalName(shapeNode, "spPr");
        const prstGeom = getDirectChildByLocalName(spPr, "prstGeom");
        if (String((prstGeom === null || prstGeom === void 0 ? void 0 : prstGeom.getAttribute("prst")) || "").trim() === "rect") {
            return "rect";
        }
        return null;
    }
    function parseShapeDimensions(anchor, shapeNode) {
        const extNode = getDirectChildByLocalName(anchor, "ext")
            || getDirectChildByLocalName(getDirectChildByLocalName(getDirectChildByLocalName(shapeNode || anchor, "spPr"), "xfrm"), "ext");
        const widthEmu = Number((extNode === null || extNode === void 0 ? void 0 : extNode.getAttribute("cx")) || "");
        const heightEmu = Number((extNode === null || extNode === void 0 ? void 0 : extNode.getAttribute("cy")) || "");
        return {
            widthPx: emuToPx(widthEmu, 160),
            heightPx: emuToPx(heightEmu, 48)
        };
    }
    function renderRectLikeSvg(shapeNode, anchor, text, treatAsTextbox) {
        const { widthPx, heightPx } = parseShapeDimensions(anchor, shapeNode);
        const spPr = getDirectChildByLocalName(shapeNode, "spPr");
        const fillColor = parseHexColor(getDirectChildByLocalName(spPr, "solidFill")) || (treatAsTextbox ? "#FFFFFF" : "#F3F3F3");
        const lineNode = getDirectChildByLocalName(spPr, "ln");
        const strokeColor = parseHexColor(lineNode) || "#333333";
        const strokeWidth = Math.max(1, Math.round(Number((lineNode === null || lineNode === void 0 ? void 0 : lineNode.getAttribute("w")) || "") / 9525) || 1);
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
    function renderConnectorSvg(shapeNode, anchor) {
        const { widthPx, heightPx } = parseShapeDimensions(anchor, shapeNode);
        const spPr = getDirectChildByLocalName(shapeNode, "spPr");
        const lineNode = getDirectChildByLocalName(spPr, "ln");
        const strokeColor = parseHexColor(lineNode) || "#333333";
        const strokeWidth = Math.max(1, Math.round(Number((lineNode === null || lineNode === void 0 ? void 0 : lineNode.getAttribute("w")) || "") / 9525) || 1);
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
    function renderShapeSvg(shapeNode, anchor, sheetName, shapeIndex) {
        const kind = parseShapeKind(shapeNode);
        if (!kind)
            return null;
        let svg = "";
        if (kind === "connector") {
            svg = renderConnectorSvg(shapeNode, anchor);
        }
        else {
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
    globalThis.__xlsx2mdOfficeDrawing = {
        renderShapeSvg
    };
})();
