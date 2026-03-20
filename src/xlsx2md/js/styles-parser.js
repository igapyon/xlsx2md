(() => {
    const moduleRegistry = getXlsx2mdModuleRegistry();
    const EMPTY_BORDERS = {
        top: false,
        bottom: false,
        left: false,
        right: false
    };
    const runtimeEnv = requireXlsx2mdRuntimeEnv();
    const textDecoder = new TextDecoder("utf-8");
    const BUILTIN_FORMAT_CODES = {
        0: "General",
        1: "0",
        2: "0.00",
        3: "#,##0",
        4: "#,##0.00",
        9: "0%",
        10: "0.00%",
        11: "0.00E+00",
        12: "# ?/?",
        13: "# ??/??",
        14: "yyyy/m/d",
        15: "d-mmm-yy",
        16: "d-mmm",
        17: "mmm-yy",
        18: "h:mm AM/PM",
        19: "h:mm:ss AM/PM",
        20: "h:mm",
        21: "h:mm:ss",
        22: "m/d/yy h:mm",
        45: "mm:ss",
        46: "[h]:mm:ss",
        47: "mmss.0",
        49: "@",
        56: "m月d日"
    };
    function decodeXmlText(bytes) {
        return textDecoder.decode(bytes);
    }
    function xmlToDocument(xmlText) {
        return runtimeEnv.xmlToDocument(xmlText);
    }
    function hasBorderSide(side) {
        if (!side)
            return false;
        return side.hasAttribute("style") || side.children.length > 0;
    }
    function parseCellStyles(files) {
        const stylesBytes = files.get("xl/styles.xml");
        if (!stylesBytes) {
            return [{
                    borders: EMPTY_BORDERS,
                    numFmtId: 0,
                    formatCode: "General"
                }];
        }
        const doc = xmlToDocument(decodeXmlText(stylesBytes));
        const borderElements = Array.from(doc.getElementsByTagName("border"));
        const borders = borderElements.map((borderElement) => {
            const top = borderElement.getElementsByTagName("top")[0] || null;
            const bottom = borderElement.getElementsByTagName("bottom")[0] || null;
            const left = borderElement.getElementsByTagName("left")[0] || null;
            const right = borderElement.getElementsByTagName("right")[0] || null;
            return {
                top: hasBorderSide(top),
                bottom: hasBorderSide(bottom),
                left: hasBorderSide(left),
                right: hasBorderSide(right)
            };
        });
        const numFmtMap = new Map();
        const numFmtParent = doc.getElementsByTagName("numFmts")[0];
        if (numFmtParent) {
            for (const numFmtElement of Array.from(numFmtParent.getElementsByTagName("numFmt"))) {
                const numFmtId = Number(numFmtElement.getAttribute("numFmtId") || 0);
                const formatCode = numFmtElement.getAttribute("formatCode") || "";
                if (!Number.isNaN(numFmtId) && formatCode) {
                    numFmtMap.set(numFmtId, formatCode);
                }
            }
        }
        const xfsParent = doc.getElementsByTagName("cellXfs")[0];
        if (!xfsParent) {
            return [{
                    borders: borders[0] || EMPTY_BORDERS,
                    numFmtId: 0,
                    formatCode: "General"
                }];
        }
        const xfElements = Array.from(xfsParent.getElementsByTagName("xf"));
        const styles = xfElements.map((xfElement) => {
            const borderId = Number(xfElement.getAttribute("borderId") || 0);
            const numFmtId = Number(xfElement.getAttribute("numFmtId") || 0);
            return {
                borders: borders[borderId] || EMPTY_BORDERS,
                numFmtId,
                formatCode: numFmtMap.get(numFmtId) || BUILTIN_FORMAT_CODES[numFmtId] || "General"
            };
        });
        return styles.length > 0 ? styles : [{
                borders: EMPTY_BORDERS,
                numFmtId: 0,
                formatCode: "General"
            }];
    }
    const stylesParserApi = {
        EMPTY_BORDERS,
        BUILTIN_FORMAT_CODES,
        hasBorderSide,
        parseCellStyles
    };
    moduleRegistry.registerModule("stylesParser", stylesParserApi);
})();
