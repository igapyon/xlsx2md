(() => {
    const EMPTY_BORDERS = {
        top: false,
        bottom: false,
        left: false,
        right: false
    };
    const textDecoder = new TextDecoder("utf-8");
    const drawingHelper = globalThis.__xlsx2mdOfficeDrawing || null;
    const narrativeStructureHelper = globalThis.__xlsx2mdNarrativeStructure;
    if (!narrativeStructureHelper) {
        throw new Error("xlsx2md narrative structure module is not loaded");
    }
    const tableDetectorHelper = globalThis.__xlsx2mdTableDetector;
    if (!tableDetectorHelper) {
        throw new Error("xlsx2md table detector module is not loaded");
    }
    const markdownExportHelper = globalThis.__xlsx2mdMarkdownExport;
    if (!markdownExportHelper) {
        throw new Error("xlsx2md markdown export module is not loaded");
    }
    const stylesParserHelper = globalThis.__xlsx2mdStylesParser;
    if (!stylesParserHelper) {
        throw new Error("xlsx2md styles parser module is not loaded");
    }
    const sharedStringsHelper = globalThis.__xlsx2mdSharedStrings;
    if (!sharedStringsHelper) {
        throw new Error("xlsx2md shared strings module is not loaded");
    }
    const worksheetTablesHelper = globalThis.__xlsx2mdWorksheetTables;
    if (!worksheetTablesHelper) {
        throw new Error("xlsx2md worksheet tables module is not loaded");
    }
    const zipIoHelper = globalThis.__xlsx2mdZipIo;
    if (!zipIoHelper) {
        throw new Error("xlsx2md zip io module is not loaded");
    }
    let resolveDefinedNameScalarValue = null;
    let resolveDefinedNameRangeRef = null;
    let resolveStructuredRangeRef = null;
    const DEFAULT_CELL_WIDTH_EMU = 609600;
    const DEFAULT_CELL_HEIGHT_EMU = 190500;
    const SHAPE_BLOCK_GAP_X_EMU = DEFAULT_CELL_WIDTH_EMU * 4;
    const SHAPE_BLOCK_GAP_Y_EMU = DEFAULT_CELL_HEIGHT_EMU * 6;
    function colToLetters(col) {
        let current = col;
        let result = "";
        while (current > 0) {
            const remainder = (current - 1) % 26;
            result = String.fromCharCode(65 + remainder) + result;
            current = Math.floor((current - 1) / 26);
        }
        return result;
    }
    function lettersToCol(letters) {
        let result = 0;
        for (const ch of letters.toUpperCase()) {
            result = result * 26 + (ch.charCodeAt(0) - 64);
        }
        return result;
    }
    function parseCellAddress(address) {
        const normalized = String(address || "").trim().replace(/\$/g, "");
        const match = normalized.match(/^([A-Z]+)(\d+)$/i);
        if (!match) {
            return { row: 0, col: 0 };
        }
        return {
            col: lettersToCol(match[1]),
            row: Number(match[2])
        };
    }
    function normalizeFormulaAddress(address) {
        return String(address || "").trim().replace(/\$/g, "").toUpperCase();
    }
    function formatRange(startRow, startCol, endRow, endCol) {
        return `${colToLetters(startCol)}${startRow}-${colToLetters(endCol)}${endRow}`;
    }
    function parseRangeRef(ref) {
        const parts = String(ref || "").split(":");
        const start = parseCellAddress(parts[0] || "");
        const end = parseCellAddress(parts[1] || parts[0] || "");
        return {
            startRow: start.row,
            startCol: start.col,
            endRow: end.row,
            endCol: end.col,
            ref: ref
        };
    }
    function xmlToDocument(xmlText) {
        return new DOMParser().parseFromString(xmlText, "application/xml");
    }
    function getElementsByLocalName(root, localName) {
        const elements = Array.from(root.getElementsByTagName("*"));
        return elements.filter((element) => element.localName === localName);
    }
    function getFirstChildByLocalName(root, localName) {
        return getElementsByLocalName(root, localName)[0] || null;
    }
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
    function decodeXmlText(bytes) {
        return textDecoder.decode(bytes);
    }
    function getTextContent(node) {
        return ((node === null || node === void 0 ? void 0 : node.textContent) || "").replace(/\r\n/g, "\n");
    }
    function normalizeZipPath(baseFilePath, targetPath) {
        const baseDirParts = baseFilePath.split("/").slice(0, -1);
        const inputParts = targetPath.split("/");
        const parts = targetPath.startsWith("/") ? [] : baseDirParts;
        for (const part of inputParts) {
            if (!part || part === ".")
                continue;
            if (part === "..") {
                parts.pop();
            }
            else {
                parts.push(part);
            }
        }
        return parts.join("/");
    }
    function isDateFormatCode(formatCode) {
        const normalized = String(formatCode || "")
            .toLowerCase()
            .replace(/\[[^\]]*]/g, "")
            .replace(/"[^"]*"/g, "")
            .replace(/\\./g, "");
        if (!normalized)
            return false;
        if (normalized.includes("general"))
            return false;
        return /[ymdhs]/.test(normalized);
    }
    function normalizeNumericFormatCode(formatCode) {
        return String(formatCode || "")
            .trim()
            .replace(/\[[^\]]*]/g, "")
            .replace(/"([^"]*)"/g, "$1")
            .replace(/\\(.)/g, "$1")
            .replace(/_.?/g, "")
            .replace(/\*/g, "");
    }
    function excelSerialToIsoText(serial) {
        if (!Number.isFinite(serial))
            return String(serial);
        const wholeDays = Math.floor(serial);
        const fractional = serial - wholeDays;
        const utcDays = wholeDays > 59 ? wholeDays - 1 : wholeDays;
        const baseUtcMs = Date.UTC(1899, 11, 31);
        const msPerDay = 24 * 60 * 60 * 1000;
        const date = new Date(baseUtcMs + utcDays * msPerDay + Math.round(fractional * msPerDay));
        const yyyy = String(date.getUTCFullYear()).padStart(4, "0");
        const mm = String(date.getUTCMonth() + 1).padStart(2, "0");
        const dd = String(date.getUTCDate()).padStart(2, "0");
        const hh = String(date.getUTCHours()).padStart(2, "0");
        const mi = String(date.getUTCMinutes()).padStart(2, "0");
        const ss = String(date.getUTCSeconds()).padStart(2, "0");
        if (hh === "00" && mi === "00" && ss === "00") {
            return `${yyyy}-${mm}-${dd}`;
        }
        return `${yyyy}-${mm}-${dd} ${hh}:${mi}:${ss}`;
    }
    function excelSerialToDateParts(serial) {
        if (!Number.isFinite(serial))
            return null;
        const wholeDays = Math.floor(serial);
        const fractional = serial - wholeDays;
        const excelEpochOffsetDays = 25569;
        const msPerDay = 24 * 60 * 60 * 1000;
        const utcDays = wholeDays - excelEpochOffsetDays;
        const baseUtcMs = Date.UTC(1970, 0, 1);
        const date = new Date(baseUtcMs + utcDays * msPerDay + Math.round(fractional * msPerDay));
        return {
            year: date.getUTCFullYear(),
            month: date.getUTCMonth() + 1,
            day: date.getUTCDate(),
            hour: date.getUTCHours(),
            minute: date.getUTCMinutes(),
            second: date.getUTCSeconds(),
            yyyy: String(date.getUTCFullYear()).padStart(4, "0"),
            mm: String(date.getUTCMonth() + 1).padStart(2, "0"),
            dd: String(date.getUTCDate()).padStart(2, "0"),
            hh: String(date.getUTCHours()).padStart(2, "0"),
            mi: String(date.getUTCMinutes()).padStart(2, "0"),
            ss: String(date.getUTCSeconds()).padStart(2, "0")
        };
    }
    function formatTextFunctionValue(value, formatText) {
        const format = String(formatText || "").trim();
        if (!format)
            return null;
        const numericValue = Number(value);
        const normalized = format.toLowerCase();
        if (!Number.isNaN(numericValue)) {
            if (/(^|[^a-z])yyyy/.test(normalized) || normalized.includes("hh:") || normalized.includes("mm/") || normalized.includes("mm-")) {
                const parts = excelSerialToDateParts(numericValue);
                if (!parts)
                    return null;
                if (normalized === "yyyy-mm-dd")
                    return `${parts.yyyy}-${parts.mm}-${parts.dd}`;
                if (normalized === "yyyy/mm/dd")
                    return `${parts.yyyy}/${parts.mm}/${parts.dd}`;
                if (normalized === "hh:mm:ss")
                    return `${parts.hh}:${parts.mi}:${parts.ss}`;
                if (normalized === "yyyy-mm-dd hh:mm:ss")
                    return `${parts.yyyy}-${parts.mm}-${parts.dd} ${parts.hh}:${parts.mi}:${parts.ss}`;
            }
            if (/^0(?:\.0+)?$/.test(format)) {
                const decimalPlaces = (format.split(".")[1] || "").length;
                return numericValue.toFixed(decimalPlaces);
            }
            if (/^#,##0(?:\.0+)?$/.test(format)) {
                const decimalPlaces = (format.split(".")[1] || "").length;
                return numericValue.toLocaleString("en-US", {
                    minimumFractionDigits: decimalPlaces,
                    maximumFractionDigits: decimalPlaces,
                    useGrouping: true
                });
            }
        }
        return null;
    }
    function formatNumberByPattern(value, pattern) {
        const normalizedPattern = pattern.trim();
        const decimalPlaces = (normalizedPattern.split(".")[1] || "").replace(/[^0#]/g, "").length;
        const useGrouping = normalizedPattern.includes(",");
        return value.toLocaleString("en-US", {
            minimumFractionDigits: decimalPlaces,
            maximumFractionDigits: decimalPlaces,
            useGrouping
        });
    }
    function formatDateByPattern(parts, formatCode) {
        const normalized = normalizeNumericFormatCode(formatCode).toLowerCase();
        if (normalized === "yyyy/m/d") {
            return `${parts.year}/${parts.month}/${parts.day}`;
        }
        if (normalized === "m月d日") {
            return `${parts.month}月${parts.day}日`;
        }
        if (normalized === "yyyy-mm-dd") {
            return `${parts.yyyy}-${parts.mm}-${parts.dd}`;
        }
        if (normalized === "yyyy/mm/dd") {
            return `${parts.year}/${parts.month}/${parts.day}`;
        }
        if (normalized === "hh:mm:ss") {
            return `${parts.hh}:${parts.mi}:${parts.ss}`;
        }
        if (normalized.includes("ggge年m月d日")) {
            if (parts.year >= 2019) {
                const reiwaYear = parts.year - 2018;
                return `令和${reiwaYear}年${parts.month}月${parts.day}日`;
            }
            if (parts.year >= 1989) {
                const heiseiYear = parts.year - 1988;
                return `平成${heiseiYear}年${parts.month}月${parts.day}日`;
            }
            return `${parts.year}年${parts.month}月${parts.day}日`;
        }
        return null;
    }
    function formatFractionPattern(value) {
        if (!Number.isFinite(value))
            return null;
        const tolerance = 1e-9;
        for (let denominator = 1; denominator <= 100; denominator += 1) {
            const numerator = Math.round(value * denominator);
            if (Math.abs(value - (numerator / denominator)) < tolerance) {
                return `${numerator}/${denominator}`;
            }
        }
        return null;
    }
    function formatDbNum3Pattern(rawValue) {
        return rawValue.split("").join(" ");
    }
    function splitFormatSections(formatCode) {
        const sections = [];
        let current = "";
        let inQuotes = false;
        for (let index = 0; index < formatCode.length; index += 1) {
            const char = formatCode[index];
            if (char === "\"") {
                inQuotes = !inQuotes;
                current += char;
                continue;
            }
            if (char === ";" && !inQuotes) {
                sections.push(current);
                current = "";
                continue;
            }
            current += char;
        }
        sections.push(current);
        return sections;
    }
    function formatZeroSection(section) {
        const normalizedSection = String(section || "");
        if (!normalizedSection)
            return null;
        const compact = normalizedSection.replace(/_.|\\.|[*?]/g, "").trim();
        const hasDashLiteral = /"-"|(^|[^a-z0-9])-($|[^a-z0-9])/i.test(compact);
        if (!hasDashLiteral)
            return null;
        if (compact.includes("¥"))
            return "¥ -";
        if (compact.includes("$"))
            return "$ -";
        return "-";
    }
    function formatCellDisplayValue(rawValue, cellStyle) {
        var _a;
        if (rawValue === "")
            return null;
        const numericValue = Number(rawValue);
        const formatCode = normalizeNumericFormatCode(cellStyle.formatCode);
        const normalized = formatCode.toLowerCase();
        const formatSections = splitFormatSections(formatCode);
        if (!Number.isNaN(numericValue) && isDateFormatCode(formatCode)) {
            const parts = excelSerialToDateParts(numericValue);
            if (!parts)
                return null;
            const directFormatted = formatDateByPattern(parts, formatCode);
            if (directFormatted !== null) {
                return directFormatted;
            }
            const hasDate = /y/.test(normalized)
                || /d/.test(normalized)
                || /(^|[^a-z])m(?:\/|-)/.test(normalized)
                || /(?:\/|-)m(?:[^a-z]|$)/.test(normalized);
            const hasTime = /h/.test(normalized) || /s/.test(normalized) || normalized.includes(":") || normalized.includes("am/pm");
            if (hasDate && hasTime) {
                return `${parts.yyyy}-${parts.mm}-${parts.dd} ${parts.hh}:${parts.mi}:${parts.ss}`;
            }
            if (hasTime && !hasDate) {
                return `${parts.hh}:${parts.mi}:${parts.ss}`;
            }
            return `${parts.yyyy}-${parts.mm}-${parts.dd}`;
        }
        if (Number.isNaN(numericValue)) {
            return null;
        }
        if (numericValue === 0 && formatSections[2]) {
            const zeroText = formatZeroSection(formatSections[2]);
            if (zeroText) {
                return zeroText;
            }
        }
        if (normalized.includes("%")) {
            const percentPattern = normalized.split(";")[0] || normalized;
            const decimalPlaces = (percentPattern.split(".")[1] || "").replace(/[^0#]/g, "").length;
            return `${(numericValue * 100).toFixed(decimalPlaces)}%`;
        }
        if (cellStyle.numFmtId === 186 || /dbnum3/i.test(formatCode)) {
            return formatDbNum3Pattern(rawValue);
        }
        if (cellStyle.numFmtId === 42) {
            return `¥ ${formatNumberByPattern(numericValue, "#,##0").replace(/^-/, "")}`;
        }
        if (/[#0][^;]*e\+0+/i.test(formatCode)) {
            const scientificPattern = formatCode.split(";")[0] || formatCode;
            const decimalPartMatch = scientificPattern.match(/\.([0#]+)e\+/i);
            const decimalPlaces = ((decimalPartMatch === null || decimalPartMatch === void 0 ? void 0 : decimalPartMatch[1]) || "").length;
            const exponentDigits = (((_a = scientificPattern.match(/e\+([0#]+)/i)) === null || _a === void 0 ? void 0 : _a[1]) || "").length;
            const [mantissa, exponentPart] = numericValue.toExponential(decimalPlaces).split("e");
            const exponent = Number(exponentPart || 0);
            const sign = exponent >= 0 ? "+" : "-";
            const paddedExponent = String(Math.abs(exponent)).padStart(exponentDigits, "0");
            return `${mantissa}E${sign}${paddedExponent}`;
        }
        if (normalized.includes("?/?")) {
            return formatFractionPattern(numericValue);
        }
        if (/^[^;]*[#0,]+(?:\.[#0]+)?/.test(formatCode)) {
            const primaryPattern = (formatCode.split(";")[0] || formatCode).trim();
            if (primaryPattern.includes("¥")) {
                const numericText = formatNumberByPattern(numericValue, primaryPattern.replace(/[^#0,.\-]/g, ""));
                const withCurrency = primaryPattern.includes("*") ? `¥ ${numericText.replace(/^-/, "")}` : `¥${numericText.replace(/^-/, "")}`;
                return `${numericValue < 0 ? "-" : ""}${withCurrency}`;
            }
            return formatNumberByPattern(numericValue, primaryPattern.replace(/[^#0,.\-]/g, ""));
        }
        return null;
    }
    function applyResolvedFormulaValue(cell, resolvedValue, resolutionSource = "legacy_resolver") {
        const rawValue = String(resolvedValue || "");
        const formattedValue = formatCellDisplayValue(rawValue, {
            borders: cell.borders,
            numFmtId: cell.numFmtId,
            formatCode: cell.formatCode
        });
        cell.rawValue = rawValue;
        cell.outputValue = formattedValue !== null && formattedValue !== void 0 ? formattedValue : rawValue;
        cell.resolutionStatus = "resolved";
        cell.resolutionSource = resolutionSource;
    }
    function parseDateLikeParts(value) {
        const trimmed = String(value || "").trim();
        const numericValue = Number(trimmed);
        if (!Number.isNaN(numericValue)) {
            return excelSerialToDateParts(numericValue);
        }
        const isoMatch = trimmed.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:[ T](\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?$/);
        if (isoMatch) {
            return {
                yyyy: isoMatch[1],
                mm: isoMatch[2].padStart(2, "0"),
                dd: isoMatch[3].padStart(2, "0"),
                hh: (isoMatch[4] || "00").padStart(2, "0"),
                mi: (isoMatch[5] || "00").padStart(2, "0"),
                ss: (isoMatch[6] || "00").padStart(2, "0")
            };
        }
        const japaneseMatch = trimmed.match(/^(\d{4})年(\d{1,2})月(\d{1,2})日(?:\s*(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?$/);
        if (japaneseMatch) {
            return {
                yyyy: japaneseMatch[1],
                mm: japaneseMatch[2].padStart(2, "0"),
                dd: japaneseMatch[3].padStart(2, "0"),
                hh: (japaneseMatch[4] || "00").padStart(2, "0"),
                mi: (japaneseMatch[5] || "00").padStart(2, "0"),
                ss: (japaneseMatch[6] || "00").padStart(2, "0")
            };
        }
        const japaneseYearMonthMatch = trimmed.match(/^(\d{4})年(\d{1,2})月$/);
        if (japaneseYearMonthMatch) {
            return {
                yyyy: japaneseYearMonthMatch[1],
                mm: japaneseYearMonthMatch[2].padStart(2, "0"),
                dd: "01",
                hh: "00",
                mi: "00",
                ss: "00"
            };
        }
        const japaneseMonthDayMatch = trimmed.match(/^(\d{1,2})月(\d{1,2})日$/);
        if (japaneseMonthDayMatch) {
            return {
                yyyy: "2000",
                mm: japaneseMonthDayMatch[1].padStart(2, "0"),
                dd: japaneseMonthDayMatch[2].padStart(2, "0"),
                hh: "00",
                mi: "00",
                ss: "00"
            };
        }
        const isoYearMonthMatch = trimmed.match(/^(\d{4})[-/](\d{1,2})$/);
        if (isoYearMonthMatch) {
            return {
                yyyy: isoYearMonthMatch[1],
                mm: isoYearMonthMatch[2].padStart(2, "0"),
                dd: "01",
                hh: "00",
                mi: "00",
                ss: "00"
            };
        }
        return null;
    }
    function datePartsToExcelSerial(year, month, day, hour = 0, minute = 0, second = 0) {
        if (![year, month, day, hour, minute, second].every(Number.isFinite))
            return null;
        const baseUtcMs = Date.UTC(1899, 11, 31);
        const targetUtcMs = Date.UTC(year, month - 1, day, hour, minute, second);
        const msPerDay = 24 * 60 * 60 * 1000;
        let serial = (targetUtcMs - baseUtcMs) / msPerDay;
        if (serial >= 60) {
            serial += 1;
        }
        return serial;
    }
    function parseValueFunctionText(value) {
        const trimmed = String(value || "").trim();
        if (!trimmed)
            return null;
        const numericValue = Number(trimmed.replace(/,/g, ""));
        if (!Number.isNaN(numericValue)) {
            return numericValue;
        }
        const parts = parseDateLikeParts(trimmed);
        if (!parts)
            return null;
        return datePartsToExcelSerial(Number(parts.yyyy), Number(parts.mm), Number(parts.dd), Number(parts.hh), Number(parts.mi), Number(parts.ss));
    }
    function parseRelationships(files, relsPath, sourcePath) {
        const relBytes = files.get(relsPath);
        const relations = new Map();
        if (!relBytes) {
            return relations;
        }
        const doc = xmlToDocument(decodeXmlText(relBytes));
        const nodes = Array.from(doc.getElementsByTagName("Relationship"));
        for (const node of nodes) {
            const id = node.getAttribute("Id") || "";
            const target = node.getAttribute("Target") || "";
            if (!id || !target)
                continue;
            relations.set(id, normalizeZipPath(sourcePath, target));
        }
        return relations;
    }
    function buildRelsPath(sourcePath) {
        const parts = sourcePath.split("/");
        const fileName = parts.pop() || "";
        const dir = parts.join("/");
        return `${dir}/_rels/${fileName}.rels`;
    }
    function getImageExtension(mediaPath) {
        const match = mediaPath.match(/\.([a-z0-9]+)$/i);
        return match ? match[1].toLowerCase() : "bin";
    }
    function createSafeSheetAssetDir(sheetName) {
        return sheetName.replace(/[\\/:*?"<>|]+/g, "_").trim() || "Sheet";
    }
    function parseDrawingImages(files, sheetName, sheetPath) {
        const sheetRels = parseRelationships(files, buildRelsPath(sheetPath), sheetPath);
        const imageAssets = [];
        let imageCounter = 1;
        for (const drawingPath of sheetRels.values()) {
            if (!/\/drawings\/.+\.xml$/i.test(drawingPath))
                continue;
            const drawingBytes = files.get(drawingPath);
            if (!drawingBytes)
                continue;
            const drawingDoc = xmlToDocument(decodeXmlText(drawingBytes));
            const drawingRels = parseRelationships(files, buildRelsPath(drawingPath), drawingPath);
            const anchors = getElementsByLocalName(drawingDoc, "oneCellAnchor").concat(getElementsByLocalName(drawingDoc, "twoCellAnchor"));
            for (const anchor of anchors) {
                const from = getFirstChildByLocalName(anchor, "from");
                const colNode = getFirstChildByLocalName(from || anchor, "col");
                const rowNode = getFirstChildByLocalName(from || anchor, "row");
                const col = Number(getTextContent(colNode)) + 1;
                const row = Number(getTextContent(rowNode)) + 1;
                if (!Number.isFinite(col) || !Number.isFinite(row) || col <= 0 || row <= 0) {
                    continue;
                }
                const blip = getElementsByLocalName(anchor, "blip")[0] || null;
                const embedId = (blip === null || blip === void 0 ? void 0 : blip.getAttribute("r:embed")) || (blip === null || blip === void 0 ? void 0 : blip.getAttribute("embed")) || "";
                const mediaPath = drawingRels.get(embedId) || "";
                if (!mediaPath)
                    continue;
                const mediaBytes = files.get(mediaPath);
                if (!mediaBytes)
                    continue;
                const extension = getImageExtension(mediaPath);
                const safeDir = createSafeSheetAssetDir(sheetName);
                const filename = `image_${String(imageCounter).padStart(3, "0")}.${extension}`;
                imageAssets.push({
                    sheetName,
                    filename,
                    path: `assets/${safeDir}/${filename}`,
                    anchor: `${colToLetters(col)}${row}`,
                    data: new Uint8Array(mediaBytes),
                    mediaPath
                });
                imageCounter += 1;
            }
        }
        return imageAssets;
    }
    function parseChartType(chartDoc) {
        const typeMap = [
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
            .filter((entry) => getElementsByLocalName(chartDoc, entry.localName).length > 0)
            .map((entry) => entry.label);
        if (matched.length === 0)
            return "Chart";
        if (matched.length === 1)
            return matched[0];
        return `${matched.join(" + ")} (Combined)`;
    }
    function parseChartTitle(chartDoc) {
        const richText = getElementsByLocalName(chartDoc, "t")
            .map((node) => getTextContent(node))
            .filter(Boolean);
        if (richText.length > 0) {
            return richText.join("").trim();
        }
        return "";
    }
    function parseChartSeries(chartDoc) {
        const plotArea = getFirstChildByLocalName(chartDoc, "plotArea") || chartDoc.documentElement;
        const axisPositionById = new Map();
        for (const axisNode of getElementsByLocalName(plotArea, "valAx")) {
            const axisIdNode = getFirstChildByLocalName(axisNode, "axId");
            const axisPosNode = getFirstChildByLocalName(axisNode, "axPos");
            const axisId = (axisIdNode === null || axisIdNode === void 0 ? void 0 : axisIdNode.getAttribute("val")) || getTextContent(axisIdNode);
            const axisPos = (axisPosNode === null || axisPosNode === void 0 ? void 0 : axisPosNode.getAttribute("val")) || getTextContent(axisPosNode);
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
        const series = [];
        for (const localName of chartContainerNames) {
            for (const chartNode of getElementsByLocalName(plotArea, localName)) {
                const axisIds = getElementsByLocalName(chartNode, "axId")
                    .map((node) => node.getAttribute("val") || getTextContent(node))
                    .filter(Boolean);
                const isSecondary = axisIds.some((axisId) => axisPositionById.get(axisId) === "r");
                for (const seriesNode of getElementsByLocalName(chartNode, "ser")) {
                    const txNode = getFirstChildByLocalName(seriesNode, "tx") || seriesNode;
                    const nameRef = getFirstChildByLocalName(txNode, "f");
                    const nameValue = getFirstChildByLocalName(txNode, "v");
                    const nameText = getElementsByLocalName(txNode, "t")
                        .map((node) => getTextContent(node))
                        .join("")
                        .trim();
                    const catRef = getFirstChildByLocalName(getFirstChildByLocalName(getFirstChildByLocalName(seriesNode, "cat") || seriesNode, "strRef") || seriesNode, "f")
                        || getFirstChildByLocalName(getFirstChildByLocalName(getFirstChildByLocalName(seriesNode, "cat") || seriesNode, "numRef") || seriesNode, "f");
                    const valRef = getFirstChildByLocalName(getFirstChildByLocalName(seriesNode, "val") || seriesNode, "f")
                        || getFirstChildByLocalName(getFirstChildByLocalName(getFirstChildByLocalName(seriesNode, "val") || seriesNode, "numRef") || seriesNode, "f");
                    series.push({
                        name: nameText || getTextContent(nameValue) || getTextContent(nameRef) || "Series",
                        categoriesRef: getTextContent(catRef),
                        valuesRef: getTextContent(valRef),
                        axis: isSecondary ? "secondary" : "primary"
                    });
                }
            }
        }
        return series;
    }
    function parseDrawingCharts(files, sheetName, sheetPath) {
        const sheetRels = parseRelationships(files, buildRelsPath(sheetPath), sheetPath);
        const charts = [];
        for (const drawingPath of sheetRels.values()) {
            if (!/\/drawings\/.+\.xml$/i.test(drawingPath))
                continue;
            const drawingBytes = files.get(drawingPath);
            if (!drawingBytes)
                continue;
            const drawingDoc = xmlToDocument(decodeXmlText(drawingBytes));
            const drawingRels = parseRelationships(files, buildRelsPath(drawingPath), drawingPath);
            const anchors = getElementsByLocalName(drawingDoc, "oneCellAnchor").concat(getElementsByLocalName(drawingDoc, "twoCellAnchor"));
            for (const anchor of anchors) {
                const from = getFirstChildByLocalName(anchor, "from");
                const colNode = getFirstChildByLocalName(from || anchor, "col");
                const rowNode = getFirstChildByLocalName(from || anchor, "row");
                const col = Number(getTextContent(colNode)) + 1;
                const row = Number(getTextContent(rowNode)) + 1;
                if (!Number.isFinite(col) || !Number.isFinite(row) || col <= 0 || row <= 0) {
                    continue;
                }
                const chartNode = getFirstChildByLocalName(anchor, "graphicFrame");
                const chartRef = getElementsByLocalName(chartNode || anchor, "chart")[0] || null;
                const relId = (chartRef === null || chartRef === void 0 ? void 0 : chartRef.getAttribute("r:id")) || (chartRef === null || chartRef === void 0 ? void 0 : chartRef.getAttribute("id")) || "";
                const chartPath = drawingRels.get(relId) || "";
                if (!chartPath)
                    continue;
                const chartBytes = files.get(chartPath);
                if (!chartBytes)
                    continue;
                const chartDoc = xmlToDocument(decodeXmlText(chartBytes));
                charts.push({
                    sheetName,
                    anchor: `${colToLetters(col)}${row}`,
                    chartPath,
                    title: parseChartTitle(chartDoc),
                    chartType: parseChartType(chartDoc),
                    series: parseChartSeries(chartDoc)
                });
            }
        }
        return charts;
    }
    function parseShapeKind(shapeNode) {
        if (!shapeNode)
            return "Shape";
        if (shapeNode.localName === "cxnSp") {
            const geomNode = getFirstChildByLocalName(getFirstChildByLocalName(shapeNode, "spPr") || shapeNode, "prstGeom");
            const prst = String((geomNode === null || geomNode === void 0 ? void 0 : geomNode.getAttribute("prst")) || "").trim();
            if (prst === "straightConnector1") {
                return "Straight Arrow Connector";
            }
            return prst ? `Connector (${prst})` : "Connector";
        }
        if (shapeNode.localName !== "sp") {
            return "Shape";
        }
        const nvSpPr = getFirstChildByLocalName(shapeNode, "nvSpPr");
        const cNvSpPr = getFirstChildByLocalName(nvSpPr || shapeNode, "cNvSpPr");
        if ((cNvSpPr === null || cNvSpPr === void 0 ? void 0 : cNvSpPr.getAttribute("txBox")) === "1") {
            return "Text Box";
        }
        const geomNode = getFirstChildByLocalName(getFirstChildByLocalName(shapeNode, "spPr") || shapeNode, "prstGeom");
        const prst = String((geomNode === null || geomNode === void 0 ? void 0 : geomNode.getAttribute("prst")) || "").trim();
        if (prst === "rect") {
            return "Rectangle";
        }
        return prst ? `Shape (${prst})` : "Shape";
    }
    function parseShapeText(shapeNode) {
        return getElementsByLocalName(shapeNode || document, "t")
            .map((node) => getTextContent(node))
            .filter(Boolean)
            .join("")
            .trim();
    }
    function parseShapeExt(anchor, shapeNode) {
        const extNode = getDirectChildByLocalName(anchor, "ext")
            || getDirectChildByLocalName(getDirectChildByLocalName(getDirectChildByLocalName(shapeNode || anchor, "spPr") || shapeNode || anchor, "xfrm"), "ext");
        const widthEmu = Number((extNode === null || extNode === void 0 ? void 0 : extNode.getAttribute("cx")) || "");
        const heightEmu = Number((extNode === null || extNode === void 0 ? void 0 : extNode.getAttribute("cy")) || "");
        return {
            widthEmu: Number.isFinite(widthEmu) ? widthEmu : null,
            heightEmu: Number.isFinite(heightEmu) ? heightEmu : null
        };
    }
    function flattenXmlNodeEntries(node, path = "", entries = []) {
        if (!node)
            return entries;
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
                flattenXmlNodeEntries(child, currentPath, entries);
            }
        }
        return entries;
    }
    function parseShapeRawEntries(anchor) {
        const entries = [];
        return flattenXmlNodeEntries(anchor, "", entries);
    }
    function renderHierarchicalRawEntries(entries) {
        const root = {
            children: new Map(),
            value: null
        };
        for (const entry of entries) {
            const parts = entry.key.split("/").filter(Boolean);
            let current = root;
            for (const part of parts) {
                if (!current.children.has(part)) {
                    current.children.set(part, {
                        children: new Map(),
                        value: null
                    });
                }
                current = current.children.get(part);
            }
            current.value = entry.value;
        }
        const lines = [];
        function visit(node, depth) {
            for (const [key, child] of node.children.entries()) {
                const indent = " ".repeat(depth * 4);
                if (child.value !== null) {
                    lines.push(`${indent}- \`${key}\`: \`${child.value}\``);
                }
                else {
                    lines.push(`${indent}- \`${key}\``);
                }
                visit(child, depth + 1);
            }
        }
        visit(root, 0);
        return lines;
    }
    function parseAnchorInt(anchor, parentName, childName) {
        const parent = getFirstChildByLocalName(anchor || document, parentName);
        const child = getFirstChildByLocalName(parent || anchor || document, childName);
        const value = Number(getTextContent(child));
        return Number.isFinite(value) ? value : null;
    }
    function parseShapeBoundingBox(anchor, shapeNode, widthEmu, heightEmu) {
        const fromCol = parseAnchorInt(anchor, "from", "col") || 0;
        const fromRow = parseAnchorInt(anchor, "from", "row") || 0;
        const fromColOff = parseAnchorInt(anchor, "from", "colOff") || 0;
        const fromRowOff = parseAnchorInt(anchor, "from", "rowOff") || 0;
        const toCol = parseAnchorInt(anchor, "to", "col");
        const toRow = parseAnchorInt(anchor, "to", "row");
        const toColOff = parseAnchorInt(anchor, "to", "colOff") || 0;
        const toRowOff = parseAnchorInt(anchor, "to", "rowOff") || 0;
        const left = fromCol * DEFAULT_CELL_WIDTH_EMU + fromColOff;
        const top = fromRow * DEFAULT_CELL_HEIGHT_EMU + fromRowOff;
        if (toCol !== null && toRow !== null) {
            return {
                left,
                top,
                right: toCol * DEFAULT_CELL_WIDTH_EMU + toColOff,
                bottom: toRow * DEFAULT_CELL_HEIGHT_EMU + toRowOff
            };
        }
        const ext = parseShapeExt(anchor, shapeNode);
        return {
            left,
            top,
            right: left + Math.max(1, ext.widthEmu || widthEmu || DEFAULT_CELL_WIDTH_EMU),
            bottom: top + Math.max(1, ext.heightEmu || heightEmu || DEFAULT_CELL_HEIGHT_EMU)
        };
    }
    function bboxGap(a, b) {
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
    function extractShapeBlocks(shapes) {
        if (shapes.length === 0)
            return [];
        const visited = new Array(shapes.length).fill(false);
        const blocks = [];
        for (let i = 0; i < shapes.length; i += 1) {
            if (visited[i])
                continue;
            const queue = [i];
            visited[i] = true;
            const shapeIndexes = [];
            while (queue.length > 0) {
                const currentIndex = queue.shift();
                shapeIndexes.push(currentIndex);
                const current = shapes[currentIndex];
                for (let j = 0; j < shapes.length; j += 1) {
                    if (visited[j])
                        continue;
                    const other = shapes[j];
                    const { dx, dy } = bboxGap(current.bbox, other.bbox);
                    if (dx <= SHAPE_BLOCK_GAP_X_EMU && dy <= SHAPE_BLOCK_GAP_Y_EMU) {
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
                startCol: Math.floor(minLeft / DEFAULT_CELL_WIDTH_EMU) + 1,
                startRow: Math.floor(minTop / DEFAULT_CELL_HEIGHT_EMU) + 1,
                endCol: Math.floor(maxRight / DEFAULT_CELL_WIDTH_EMU) + 1,
                endRow: Math.floor(maxBottom / DEFAULT_CELL_HEIGHT_EMU) + 1,
                shapeIndexes: shapeIndexes.sort((a, b) => a - b)
            });
        }
        return blocks.sort((a, b) => (a.startRow - b.startRow) || (a.startCol - b.startCol));
    }
    function parseDrawingShapes(files, sheetName, sheetPath) {
        var _a;
        const sheetRels = parseRelationships(files, buildRelsPath(sheetPath), sheetPath);
        const shapes = [];
        let shapeCounter = 1;
        for (const drawingPath of sheetRels.values()) {
            if (!/\/drawings\/.+\.xml$/i.test(drawingPath))
                continue;
            const drawingBytes = files.get(drawingPath);
            if (!drawingBytes)
                continue;
            const drawingDoc = xmlToDocument(decodeXmlText(drawingBytes));
            const anchors = getElementsByLocalName(drawingDoc, "oneCellAnchor").concat(getElementsByLocalName(drawingDoc, "twoCellAnchor"));
            for (const anchor of anchors) {
                const from = getFirstChildByLocalName(anchor, "from");
                const colNode = getFirstChildByLocalName(from || anchor, "col");
                const rowNode = getFirstChildByLocalName(from || anchor, "row");
                const col = Number(getTextContent(colNode)) + 1;
                const row = Number(getTextContent(rowNode)) + 1;
                if (!Number.isFinite(col) || !Number.isFinite(row) || col <= 0 || row <= 0) {
                    continue;
                }
                if (getElementsByLocalName(anchor, "blip").length > 0)
                    continue;
                if (getElementsByLocalName(anchor, "chart").length > 0)
                    continue;
                const shapeNode = getFirstChildByLocalName(anchor, "sp") || getFirstChildByLocalName(anchor, "cxnSp");
                if (!shapeNode)
                    continue;
                const cNvPr = getFirstChildByLocalName(getFirstChildByLocalName(shapeNode, shapeNode.localName === "sp" ? "nvSpPr" : "nvCxnSpPr") || shapeNode, "cNvPr");
                const { widthEmu, heightEmu } = parseShapeExt(anchor, shapeNode);
                const svgAsset = ((_a = drawingHelper === null || drawingHelper === void 0 ? void 0 : drawingHelper.renderShapeSvg) === null || _a === void 0 ? void 0 : _a.call(drawingHelper, shapeNode, anchor, sheetName, shapeCounter)) || null;
                shapes.push({
                    sheetName,
                    anchor: `${colToLetters(col)}${row}`,
                    name: String((cNvPr === null || cNvPr === void 0 ? void 0 : cNvPr.getAttribute("name")) || "").trim() || "Shape",
                    kind: parseShapeKind(shapeNode),
                    text: parseShapeText(shapeNode),
                    widthEmu,
                    heightEmu,
                    elementName: `xdr:${shapeNode.localName}`,
                    anchorElementName: anchor.tagName || anchor.nodeName || anchor.localName || "anchor",
                    rawEntries: parseShapeRawEntries(anchor),
                    bbox: parseShapeBoundingBox(anchor, shapeNode, widthEmu, heightEmu),
                    svgFilename: (svgAsset === null || svgAsset === void 0 ? void 0 : svgAsset.filename) || null,
                    svgPath: (svgAsset === null || svgAsset === void 0 ? void 0 : svgAsset.path) || null,
                    svgData: (svgAsset === null || svgAsset === void 0 ? void 0 : svgAsset.data) || null
                });
                shapeCounter += 1;
            }
        }
        return shapes;
    }
    function extractCellOutputValue(cellElement, sharedStrings, cellStyle, formulaOverride = "") {
        const type = (cellElement.getAttribute("t") || "").trim();
        const valueNode = cellElement.getElementsByTagName("v")[0] || null;
        const valueText = getTextContent(valueNode);
        const formulaText = formulaOverride || getTextContent(cellElement.getElementsByTagName("f")[0]);
        const cachedValueState = !formulaText
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
                    cachedValueState
                };
            }
            if (valueNode) {
                const formattedValue = formatCellDisplayValue(valueText, cellStyle);
                return {
                    valueType: type || "formula",
                    rawValue: valueText,
                    outputValue: formattedValue !== null && formattedValue !== void 0 ? formattedValue : valueText,
                    formulaText: normalizedFormula,
                    resolutionStatus: "resolved",
                    resolutionSource: "cached_value",
                    cachedValueState
                };
            }
            return {
                valueType: type || "formula",
                rawValue: normalizedFormula,
                outputValue: normalizedFormula,
                formulaText: normalizedFormula,
                resolutionStatus: "fallback_formula",
                resolutionSource: "formula_text",
                cachedValueState
            };
        }
        if (type === "s") {
            const sharedIndex = Number(valueText || 0);
            return {
                valueType: type,
                rawValue: valueText,
                outputValue: sharedStrings[sharedIndex] || "",
                formulaText: "",
                resolutionStatus: null,
                resolutionSource: null,
                cachedValueState: null
            };
        }
        if (type === "inlineStr") {
            const inlineText = Array.from(cellElement.getElementsByTagName("t")).map((node) => getTextContent(node)).join("");
            return {
                valueType: type,
                rawValue: inlineText,
                outputValue: inlineText,
                formulaText: "",
                resolutionStatus: null,
                resolutionSource: null,
                cachedValueState: null
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
                cachedValueState: null
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
                cachedValueState: null
            };
        }
        if (valueText) {
            const formattedValue = formatCellDisplayValue(valueText, cellStyle);
            if (formattedValue !== null) {
                return {
                    valueType: type,
                    rawValue: valueText,
                    outputValue: formattedValue,
                    formulaText: "",
                    resolutionStatus: null,
                    resolutionSource: null,
                    cachedValueState: null
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
            cachedValueState: null
        };
    }
    function shiftReferenceAddress(addressText, rowOffset, colOffset) {
        const match = String(addressText || "").match(/^(\$?)([A-Z]+)(\$?)(\d+)$/i);
        if (!match)
            return addressText;
        const colAbsolute = match[1] === "$";
        const rowAbsolute = match[3] === "$";
        const baseCol = lettersToCol(match[2]);
        const baseRow = Number(match[4]);
        const shiftedCol = colAbsolute ? baseCol : baseCol + colOffset;
        const shiftedRow = rowAbsolute ? baseRow : baseRow + rowOffset;
        const safeCol = Math.max(1, shiftedCol);
        const safeRow = Math.max(1, shiftedRow);
        return `${colAbsolute ? "$" : ""}${colToLetters(safeCol)}${rowAbsolute ? "$" : ""}${safeRow}`;
    }
    function translateSharedFormula(baseFormulaText, baseAddress, targetAddress) {
        const basePos = parseCellAddress(baseAddress);
        const targetPos = parseCellAddress(targetAddress);
        if (!basePos.row || !basePos.col || !targetPos.row || !targetPos.col) {
            return baseFormulaText;
        }
        const rowOffset = targetPos.row - basePos.row;
        const colOffset = targetPos.col - basePos.col;
        const normalized = String(baseFormulaText || "").replace(/^=/, "");
        const translated = normalized.replace(/(?:'((?:[^']|'')+)'|([A-Za-z0-9_ ]+))!(\$?[A-Z]+\$?\d+)|(\$?[A-Z]+\$?\d+)/g, (full, quotedSheet, plainSheet, qualifiedAddress, localAddress) => {
            const address = qualifiedAddress || localAddress;
            if (!address)
                return full;
            const shifted = shiftReferenceAddress(address, rowOffset, colOffset);
            if (qualifiedAddress) {
                const sheetPrefix = quotedSheet ? `'${quotedSheet}'` : plainSheet;
                return `${sheetPrefix}!${shifted}`;
            }
            return shifted;
        });
        return translated.startsWith("=") ? translated : `=${translated}`;
    }
    function parseSimpleFormulaReference(formulaText, currentSheetName) {
        const normalizedFormula = String(formulaText || "").trim().replace(/^=/, "");
        const quotedSheetMatch = normalizedFormula.match(/^'((?:[^']|'')+)'!(\$?[A-Z]+\$?\d+)$/i);
        if (quotedSheetMatch) {
            return {
                sheetName: quotedSheetMatch[1].replace(/''/g, "'"),
                address: normalizeFormulaAddress(quotedSheetMatch[2])
            };
        }
        const sheetMatch = normalizedFormula.match(/^([^'=][^!]*)!(\$?[A-Z]+\$?\d+)$/i);
        if (sheetMatch) {
            return {
                sheetName: sheetMatch[1],
                address: normalizeFormulaAddress(sheetMatch[2])
            };
        }
        const localMatch = normalizedFormula.match(/^(\$?[A-Z]+\$?\d+)$/i);
        if (localMatch) {
            return {
                sheetName: currentSheetName,
                address: normalizeFormulaAddress(localMatch[1])
            };
        }
        return null;
    }
    function parseSheetScopedDefinedNameReference(expression, currentSheetName) {
        const normalizedExpression = String(expression || "").trim();
        const quotedSheetMatch = normalizedExpression.match(/^'((?:[^']|'')+)'!([A-Za-z_][A-Za-z0-9_.]*)$/);
        if (quotedSheetMatch) {
            return {
                sheetName: normalizeFormulaSheetName(quotedSheetMatch[1].replace(/''/g, "'")),
                name: quotedSheetMatch[2]
            };
        }
        const sheetMatch = normalizedExpression.match(/^([^'=][^!]*)!([A-Za-z_][A-Za-z0-9_.]*)$/);
        if (!sheetMatch) {
            return null;
        }
        if (/^\$?[A-Z]+\$?\d+$/i.test(sheetMatch[2])) {
            return null;
        }
        return {
            sheetName: normalizeFormulaSheetName(sheetMatch[1] || currentSheetName),
            name: sheetMatch[2]
        };
    }
    function normalizeFormulaSheetName(rawName) {
        return String(rawName || "").replace(/^'/, "").replace(/'$/, "").replace(/''/g, "'");
    }
    function normalizeDefinedNameKey(name) {
        return String(name || "").trim().toUpperCase();
    }
    function parseRangeAddress(rawRange) {
        const match = String(rawRange || "").trim().match(/^(\$?[A-Z]+\$?\d+):(\$?[A-Z]+\$?\d+)$/i);
        if (!match)
            return null;
        return {
            start: normalizeFormulaAddress(match[1]),
            end: normalizeFormulaAddress(match[2])
        };
    }
    function parseDefinedNames(workbookDoc, sheetNames) {
        const result = [];
        const definedNameElements = Array.from(workbookDoc.getElementsByTagName("definedName"));
        for (const element of definedNameElements) {
            const name = element.getAttribute("name") || "";
            if (!name || name.startsWith("_xlnm."))
                continue;
            const formulaText = getTextContent(element).trim();
            if (!formulaText)
                continue;
            const localSheetIdText = element.getAttribute("localSheetId");
            const localSheetId = localSheetIdText == null || localSheetIdText === "" ? Number.NaN : Number(localSheetIdText);
            result.push({
                name,
                formulaText: formulaText.startsWith("=") ? formulaText : `=${formulaText}`,
                localSheetName: Number.isNaN(localSheetId) ? null : (sheetNames[localSheetId] || null)
            });
        }
        return result;
    }
    function buildFormulaResolver(workbook) {
        const sheetMap = new Map();
        const cellMaps = new Map();
        const tableMap = new Map();
        for (const sheet of workbook.sheets) {
            sheetMap.set(sheet.name, sheet);
            const cellMap = new Map();
            for (const cell of sheet.cells) {
                cellMap.set(cell.address.toUpperCase(), cell);
            }
            cellMaps.set(sheet.name, cellMap);
            for (const table of sheet.tables) {
                if (table.name) {
                    tableMap.set(worksheetTablesHelper.normalizeStructuredTableKey(table.name), table);
                }
                if (table.displayName) {
                    tableMap.set(worksheetTablesHelper.normalizeStructuredTableKey(table.displayName), table);
                }
            }
        }
        const resolvingKeys = new Set();
        const definedNameMap = new Map();
        for (const entry of workbook.definedNames) {
            const key = entry.localSheetName
                ? `${normalizeFormulaSheetName(entry.localSheetName)}::${normalizeDefinedNameKey(entry.name)}`
                : `::${normalizeDefinedNameKey(entry.name)}`;
            definedNameMap.set(key, entry.formulaText);
        }
        function lookupDefinedNameFormula(sheetName, name) {
            const normalizedName = normalizeDefinedNameKey(name);
            return definedNameMap.get(`${normalizeFormulaSheetName(sheetName)}::${normalizedName}`)
                || definedNameMap.get(`::${normalizedName}`)
                || null;
        }
        function resolveCellValue(sheetName, address) {
            var _a;
            const sheet = sheetMap.get(sheetName);
            if (!sheet)
                return "#REF!";
            const cell = ((_a = cellMaps.get(sheetName)) === null || _a === void 0 ? void 0 : _a.get(address.toUpperCase())) || null;
            if (!cell)
                return "";
            const key = `${sheetName}!${address.toUpperCase()}`;
            if (resolvingKeys.has(key)) {
                return "";
            }
            if (cell.formulaText && (!cell.outputValue || cell.resolutionStatus !== "resolved")) {
                resolvingKeys.add(key);
                try {
                    try {
                        const result = tryResolveFormulaExpressionDetailed(cell.formulaText, sheetName, resolveCellValue, undefined, undefined, cell.address);
                        if ((result === null || result === void 0 ? void 0 : result.value) != null) {
                            applyResolvedFormulaValue(cell, result.value, result.source || "legacy_resolver");
                        }
                    }
                    catch (error) {
                        if (!(error instanceof Error) || error.message !== "__FORMULA_UNRESOLVED__") {
                            throw error;
                        }
                    }
                }
                finally {
                    resolvingKeys.delete(key);
                }
            }
            if (cell.formulaText) {
                if (cell.resolutionStatus === "resolved") {
                    const rawValue = String(cell.rawValue || "");
                    const outputValue = String(cell.outputValue || "");
                    if (rawValue && rawValue !== cell.formulaText) {
                        return rawValue;
                    }
                    return outputValue || rawValue;
                }
                const rawValue = String(cell.rawValue || "");
                const outputValue = String(cell.outputValue || "");
                if (rawValue && rawValue !== cell.formulaText) {
                    return rawValue;
                }
                if (outputValue && outputValue !== cell.formulaText) {
                    return outputValue;
                }
                return "";
            }
            if (["s", "inlineStr", "str", "e", "b"].includes(cell.valueType)) {
                return String(cell.outputValue || cell.rawValue || "");
            }
            return String(cell.rawValue || cell.outputValue || "");
        }
        function resolveDefinedNameValue(sheetName, name) {
            const formulaText = lookupDefinedNameFormula(sheetName, name);
            if (!formulaText)
                return null;
            const directRef = parseSimpleFormulaReference(formulaText, sheetName);
            if (directRef) {
                const value = resolveCellValue(directRef.sheetName, directRef.address);
                return value === "" ? null : value;
            }
            const scalar = resolveScalarFormulaValue(formulaText.replace(/^=/, ""), sheetName, resolveCellValue);
            return scalar == null || scalar === "" ? null : scalar;
        }
        function resolveDefinedNameRange(sheetName, name) {
            const formulaText = lookupDefinedNameFormula(sheetName, name);
            if (!formulaText)
                return null;
            const normalized = formulaText.replace(/^=/, "").trim();
            const directRange = parseQualifiedRangeReference(normalized, sheetName);
            if (directRange) {
                return directRange;
            }
            const separatorIndex = findTopLevelOperatorIndex(normalized, ":");
            if (separatorIndex <= 0) {
                return null;
            }
            const leftText = normalized.slice(0, separatorIndex).trim();
            const rightText = normalized.slice(separatorIndex + 1).trim();
            const startRef = parseSimpleFormulaReference(`=${leftText}`, sheetName);
            const indexCall = parseWholeFunctionCall(rightText, ["INDEX"]);
            if (!startRef || !indexCall) {
                return null;
            }
            const args = splitFormulaArguments(indexCall.argsText.trim());
            if (args.length < 2 || args.length > 3) {
                return null;
            }
            const rangeRef = parseQualifiedRangeReference(args[0], sheetName);
            const rowIndex = Number(resolveScalarFormulaValue(args[1], sheetName, resolveCellValue, (targetSheetName, rangeText) => resolveRangeEntries(targetSheetName, rangeText).numericValues, resolveRangeEntries));
            const colIndex = args.length === 3
                ? Number(resolveScalarFormulaValue(args[2], sheetName, resolveCellValue, (targetSheetName, rangeText) => resolveRangeEntries(targetSheetName, rangeText).numericValues, resolveRangeEntries))
                : 1;
            if (!rangeRef || Number.isNaN(rowIndex) || Number.isNaN(colIndex) || rowIndex < 1 || colIndex < 1) {
                return null;
            }
            const rangeStart = parseCellAddress(rangeRef.start);
            const rangeEnd = parseCellAddress(rangeRef.end);
            if (!rangeStart.row || !rangeStart.col || !rangeEnd.row || !rangeEnd.col) {
                return null;
            }
            const startRow = Math.min(rangeStart.row, rangeEnd.row);
            const endRow = Math.max(rangeStart.row, rangeEnd.row);
            const startCol = Math.min(rangeStart.col, rangeEnd.col);
            const endCol = Math.max(rangeStart.col, rangeEnd.col);
            const targetRow = startRow + Math.trunc(rowIndex) - 1;
            const targetCol = startCol + Math.trunc(colIndex) - 1;
            if (targetRow > endRow || targetCol > endCol) {
                return null;
            }
            return {
                sheetName: startRef.sheetName,
                start: startRef.address,
                end: `${colToLetters(targetCol)}${targetRow}`
            };
        }
        function resolveStructuredRange(sheetName, text) {
            const match = String(text || "").trim().match(/^(.+?)\[([^\]]+)\]$/);
            if (!match)
                return null;
            const tableKey = worksheetTablesHelper.normalizeStructuredTableKey(match[1].replace(/^'(.*)'$/, "$1"));
            const columnKey = worksheetTablesHelper.normalizeStructuredTableKey(match[2]);
            if (!tableKey || !columnKey || columnKey.startsWith("#") || columnKey.startsWith("@")) {
                return null;
            }
            const table = tableMap.get(tableKey);
            if (!table)
                return null;
            const columnIndex = table.columns.findIndex((columnName) => worksheetTablesHelper.normalizeStructuredTableKey(columnName) === columnKey);
            if (columnIndex < 0)
                return null;
            const startAddress = parseCellAddress(table.start);
            const endAddress = parseCellAddress(table.end);
            if (!startAddress.row || !startAddress.col || !endAddress.row || !endAddress.col)
                return null;
            const firstDataRow = Math.min(startAddress.row, endAddress.row) + Math.max(0, table.headerRowCount);
            const lastDataRow = Math.max(startAddress.row, endAddress.row) - Math.max(0, table.totalsRowCount);
            if (firstDataRow > lastDataRow)
                return null;
            const col = Math.min(startAddress.col, endAddress.col) + columnIndex;
            const colLetters = colToLetters(col);
            return {
                sheetName: table.sheetName || sheetName,
                start: `${colLetters}${firstDataRow}`,
                end: `${colLetters}${lastDataRow}`
            };
        }
        function resolveSpillRange(sheetName, address) {
            var _a;
            const normalizedAddress = normalizeFormulaAddress(address);
            const cell = ((_a = cellMaps.get(sheetName)) === null || _a === void 0 ? void 0 : _a.get(normalizedAddress)) || null;
            if (!cell) {
                return null;
            }
            if (cell.formulaType === "array") {
                return { sheetName, start: normalizedAddress, end: normalizedAddress };
            }
            const spillRef = String(cell.spillRef || "").trim();
            if (!spillRef) {
                return { sheetName, start: normalizedAddress, end: normalizedAddress };
            }
            const directRange = parseQualifiedRangeReference(spillRef, sheetName);
            if (directRange) {
                return directRange;
            }
            if (/^[A-Za-z]+\d+:[A-Za-z]+\d+$/.test(spillRef)) {
                const [start, end] = spillRef.split(":");
                return {
                    sheetName,
                    start: normalizeFormulaAddress(start),
                    end: normalizeFormulaAddress(end)
                };
            }
            if (/^[A-Za-z]+\d+$/.test(spillRef)) {
                const only = normalizeFormulaAddress(spillRef);
                return { sheetName, start: only, end: only };
            }
            return { sheetName, start: normalizedAddress, end: normalizedAddress };
        }
        function resolveRangeEntries(sheetName, rangeText) {
            const range = parseRangeAddress(rangeText);
            if (!range) {
                return { rawValues: [], numericValues: [] };
            }
            const start = parseCellAddress(range.start);
            const end = parseCellAddress(range.end);
            if (!start.row || !start.col || !end.row || !end.col) {
                return { rawValues: [], numericValues: [] };
            }
            const startRow = Math.min(start.row, end.row);
            const endRow = Math.max(start.row, end.row);
            const startCol = Math.min(start.col, end.col);
            const endCol = Math.max(start.col, end.col);
            const rawValues = [];
            const numericValues = [];
            for (let row = startRow; row <= endRow; row += 1) {
                for (let col = startCol; col <= endCol; col += 1) {
                    const rawValue = resolveCellValue(sheetName, `${colToLetters(col)}${row}`);
                    rawValues.push(rawValue);
                    if (String(rawValue || "").trim() === "") {
                        continue;
                    }
                    const numericValue = Number(rawValue);
                    if (!Number.isNaN(numericValue)) {
                        numericValues.push(numericValue);
                    }
                }
            }
            return { rawValues, numericValues };
        }
        return {
            resolveCellValue,
            resolveRangeValues: (sheetName, rangeText) => resolveRangeEntries(sheetName, rangeText).numericValues,
            resolveRangeEntries,
            resolveDefinedNameValue,
            resolveDefinedNameRange,
            resolveStructuredRange
        };
    }
    function tryResolveFormulaExpressionDetailed(formulaText, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries, currentAddress) {
        const normalized = String(formulaText || "").trim().replace(/^=/, "");
        if (!normalized)
            return null;
        const directDefinedNameValue = (resolveDefinedNameScalarValue === null || resolveDefinedNameScalarValue === void 0 ? void 0 : resolveDefinedNameScalarValue(currentSheetName, normalized)) || null;
        if (directDefinedNameValue != null) {
            return {
                value: directDefinedNameValue,
                source: "legacy_resolver"
            };
        }
        const astResolved = tryResolveFormulaExpressionWithAst(normalized, currentSheetName, resolveCellValue, resolveRangeEntries, currentAddress);
        if (astResolved != null) {
            return {
                value: astResolved,
                source: "ast_evaluator"
            };
        }
        const legacyResolved = tryResolveFormulaExpressionLegacy(normalized, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
        if (legacyResolved == null) {
            return null;
        }
        return {
            value: legacyResolved,
            source: "legacy_resolver"
        };
    }
    function tryResolveFormulaExpression(formulaText, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries, currentAddress) {
        var _a, _b;
        return (_b = (_a = tryResolveFormulaExpressionDetailed(formulaText, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries, currentAddress)) === null || _a === void 0 ? void 0 : _a.value) !== null && _b !== void 0 ? _b : null;
    }
    function tryResolveFormulaExpressionLegacy(normalized, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries) {
        const ifResult = tryResolveIfFunction(normalized, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
        if (ifResult != null) {
            return ifResult;
        }
        const ifErrorResult = tryResolveIfErrorFunction(normalized, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
        if (ifErrorResult != null) {
            return ifErrorResult;
        }
        const logicalResult = tryResolveLogicalFunction(normalized, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
        if (logicalResult != null) {
            return logicalResult;
        }
        const concatResult = tryResolveConcatenationExpression(normalized, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
        if (concatResult != null) {
            return concatResult;
        }
        const numericFunctionResult = tryResolveNumericFunction(normalized, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
        if (numericFunctionResult != null) {
            return numericFunctionResult;
        }
        const datePartFunctionResult = tryResolveDatePartFunction(normalized, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
        if (datePartFunctionResult != null) {
            return datePartFunctionResult;
        }
        const predicateFunctionResult = tryResolvePredicateFunction(normalized, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
        if (predicateFunctionResult != null) {
            return predicateFunctionResult;
        }
        const chooseFunctionResult = tryResolveChooseFunction(normalized, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
        if (chooseFunctionResult != null) {
            return chooseFunctionResult;
        }
        const textFunctionResult = tryResolveTextFunction(normalized, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
        if (textFunctionResult != null) {
            return textFunctionResult;
        }
        const lookupFunctionResult = tryResolveLookupFunction(normalized, currentSheetName, resolveCellValue);
        if (lookupFunctionResult != null) {
            return lookupFunctionResult;
        }
        const stringFunctionResult = tryResolveStringFunction(normalized, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
        if (stringFunctionResult != null) {
            return stringFunctionResult;
        }
        const conditionalAggregateResult = tryResolveConditionalAggregateFunction(normalized, currentSheetName, resolveCellValue);
        if (conditionalAggregateResult != null) {
            return conditionalAggregateResult;
        }
        const functionResult = tryResolveAggregateFunction(normalized, currentSheetName, resolveRangeValues, resolveRangeEntries);
        if (functionResult != null) {
            return functionResult;
        }
        const comparisonResult = tryResolveComparisonExpression(normalized, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
        if (comparisonResult != null) {
            return comparisonResult;
        }
        if (/:/.test(normalized)) {
            return null;
        }
        const replacedRefs = normalized.replace(/(?:'((?:[^']|'')+)'|([A-Za-z0-9_ ]+))!(\$?[A-Z]+\$?\d+)|(\$?[A-Z]+\$?\d+)/g, (_full, quotedSheet, plainSheet, qualifiedAddress, localAddress) => {
            const sheetName = qualifiedAddress
                ? normalizeFormulaSheetName(quotedSheet || plainSheet || currentSheetName)
                : currentSheetName;
            const address = normalizeFormulaAddress(qualifiedAddress || localAddress || "");
            const rawValue = resolveCellValue(sheetName, address);
            const numericValue = Number(rawValue);
            if (rawValue === "" || Number.isNaN(numericValue)) {
                throw new Error("__FORMULA_UNRESOLVED__");
            }
            return String(numericValue);
        });
        const replaced = replaceNumericDefinedNames(replacedRefs, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
        const replacedFunctions = replaceEmbeddedNumericFunctions(replaced, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
        if (!/^[0-9+\-*/().\s]+$/.test(replacedFunctions)) {
            return null;
        }
        try {
            const value = evaluateArithmeticExpression(replacedFunctions);
            if (!Number.isFinite(value)) {
                return null;
            }
            const rounded = Math.abs(value - Math.round(value)) < 1e-10 ? Math.round(value) : value;
            return String(rounded);
        }
        catch (error) {
            if (error instanceof Error && error.message === "__FORMULA_UNRESOLVED__") {
                return null;
            }
            return null;
        }
    }
    function tryResolveFormulaExpressionWithAst(expression, currentSheetName, resolveCellValue, resolveRangeEntries, currentAddress) {
        const formulaApi = globalThis.__xlsx2mdFormula;
        if (!(formulaApi === null || formulaApi === void 0 ? void 0 : formulaApi.parseFormula) || !(formulaApi === null || formulaApi === void 0 ? void 0 : formulaApi.evaluateFormulaAst)) {
            return null;
        }
        try {
            const ast = formulaApi.parseFormula(`=${expression}`);
            const evaluated = formulaApi.evaluateFormulaAst(ast, {
                resolveCell(ref, sheet) {
                    return coerceFormulaAstScalar(resolveCellValue(sheet || currentSheetName, normalizeFormulaAddress(ref)));
                },
                resolveName(name) {
                    var _a, _b, _c;
                    const scopedRef = parseSheetScopedDefinedNameReference(name, currentSheetName);
                    if (scopedRef) {
                        const scopedValue = (_a = resolveDefinedNameScalarValue === null || resolveDefinedNameScalarValue === void 0 ? void 0 : resolveDefinedNameScalarValue(scopedRef.sheetName, scopedRef.name)) !== null && _a !== void 0 ? _a : null;
                        if (scopedValue != null) {
                            return coerceFormulaAstScalar(scopedValue);
                        }
                    }
                    const scalarValue = (_b = resolveDefinedNameScalarValue === null || resolveDefinedNameScalarValue === void 0 ? void 0 : resolveDefinedNameScalarValue(currentSheetName, name)) !== null && _b !== void 0 ? _b : null;
                    if (scalarValue != null) {
                        return coerceFormulaAstScalar(scalarValue);
                    }
                    const rangeRef = (_c = resolveDefinedNameRangeRef === null || resolveDefinedNameRangeRef === void 0 ? void 0 : resolveDefinedNameRangeRef(currentSheetName, name)) !== null && _c !== void 0 ? _c : null;
                    if (rangeRef && resolveRangeEntries) {
                        return createFormulaAstRangeMatrix(rangeRef.sheetName, rangeRef.start, rangeRef.end, resolveRangeEntries);
                    }
                    return null;
                },
                resolveScopedName(sheet, name) {
                    var _a, _b;
                    const scopedValue = (_a = resolveDefinedNameScalarValue === null || resolveDefinedNameScalarValue === void 0 ? void 0 : resolveDefinedNameScalarValue(sheet, name)) !== null && _a !== void 0 ? _a : null;
                    if (scopedValue != null) {
                        return coerceFormulaAstScalar(scopedValue);
                    }
                    const rangeRef = (_b = resolveDefinedNameRangeRef === null || resolveDefinedNameRangeRef === void 0 ? void 0 : resolveDefinedNameRangeRef(sheet, name)) !== null && _b !== void 0 ? _b : null;
                    if (rangeRef && resolveRangeEntries) {
                        return createFormulaAstRangeMatrix(rangeRef.sheetName, rangeRef.start, rangeRef.end, resolveRangeEntries);
                    }
                    return null;
                },
                resolveStructuredRef(table, column) {
                    var _a;
                    const rangeRef = (_a = resolveStructuredRangeRef === null || resolveStructuredRangeRef === void 0 ? void 0 : resolveStructuredRangeRef(currentSheetName, `${table}[${column}]`)) !== null && _a !== void 0 ? _a : null;
                    if (!rangeRef || !resolveRangeEntries) {
                        return null;
                    }
                    return createFormulaAstRangeMatrix(rangeRef.sheetName, rangeRef.start, rangeRef.end, resolveRangeEntries);
                },
                resolveRange(startRef, endRef, sheet) {
                    if (!resolveRangeEntries) {
                        return [];
                    }
                    return createFormulaAstRangeMatrix(sheet || currentSheetName, normalizeFormulaAddress(startRef), normalizeFormulaAddress(endRef), resolveRangeEntries);
                },
                resolveSpill(ref, sheet) {
                    if (!resolveRangeEntries) {
                        return [];
                    }
                    const spillRange = resolveSpillRange(sheet || currentSheetName, ref);
                    if (!spillRange) {
                        return [];
                    }
                    return createFormulaAstRangeMatrix(spillRange.sheetName, spillRange.start, spillRange.end, resolveRangeEntries);
                },
                currentCellRef: currentAddress ? normalizeFormulaAddress(currentAddress) : undefined
            });
            return serializeFormulaAstResult(evaluated);
        }
        catch (_error) {
            return null;
        }
    }
    function coerceFormulaAstScalar(value) {
        const trimmed = String(value || "").trim();
        if (!trimmed) {
            return "";
        }
        if (trimmed === "TRUE") {
            return true;
        }
        if (trimmed === "FALSE") {
            return false;
        }
        const numeric = Number(trimmed.replace(/,/g, ""));
        if (!Number.isNaN(numeric)) {
            return numeric;
        }
        return trimmed;
    }
    function createFormulaAstRangeMatrix(sheetName, startAddress, endAddress, resolveRangeEntries) {
        const range = parseRangeAddress(`${normalizeFormulaAddress(startAddress)}:${normalizeFormulaAddress(endAddress)}`);
        if (!range) {
            return [];
        }
        const start = parseCellAddress(range.start);
        const end = parseCellAddress(range.end);
        if (!start.row || !start.col || !end.row || !end.col) {
            return [];
        }
        const startRow = Math.min(start.row, end.row);
        const endRow = Math.max(start.row, end.row);
        const startCol = Math.min(start.col, end.col);
        const endCol = Math.max(start.col, end.col);
        const entries = resolveRangeEntries(sheetName, `${range.start}:${range.end}`).rawValues;
        const matrix = [];
        let index = 0;
        for (let row = startRow; row <= endRow; row += 1) {
            const rowValues = [];
            for (let col = startCol; col <= endCol; col += 1) {
                rowValues.push(coerceFormulaAstScalar(entries[index] || ""));
                index += 1;
            }
            matrix.push(rowValues);
        }
        return matrix;
    }
    function serializeFormulaAstResult(value) {
        if (value == null) {
            return null;
        }
        if (Array.isArray(value)) {
            return null;
        }
        if (typeof value === "boolean") {
            return value ? "TRUE" : "FALSE";
        }
        return String(value);
    }
    function tryResolveIfFunction(normalizedFormula, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries) {
        const call = parseWholeFunctionCall(normalizedFormula, ["IF"]);
        if (!call)
            return null;
        const args = splitFormulaArguments(call.argsText.trim());
        if (args.length !== 3)
            return null;
        const condition = evaluateFormulaCondition(args[0], currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
        if (condition == null)
            return null;
        return resolveScalarFormulaValue(condition ? args[1] : args[2], currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
    }
    function tryResolveIfErrorFunction(normalizedFormula, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries) {
        const call = parseWholeFunctionCall(normalizedFormula, ["IFERROR"]);
        if (!call)
            return null;
        const args = splitFormulaArguments(call.argsText.trim());
        if (args.length !== 2)
            return null;
        const primary = resolveScalarFormulaValue(args[0], currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
        if (primary != null && !/^#(?:[A-Z]+\/[A-Z]+|[A-Z]+[!?]?)/i.test(primary.trim())) {
            return primary;
        }
        return resolveScalarFormulaValue(args[1], currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
    }
    function tryResolveLogicalFunction(normalizedFormula, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries) {
        const call = parseWholeFunctionCall(normalizedFormula, ["AND", "OR", "NOT"]);
        if (!call)
            return null;
        const functionName = call.name;
        const args = splitFormulaArguments(call.argsText.trim());
        if (functionName === "NOT") {
            if (args.length !== 1)
                return null;
            const value = evaluateFormulaCondition(args[0], currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
            if (value == null)
                return null;
            return value ? "FALSE" : "TRUE";
        }
        if (args.length === 0)
            return null;
        const evaluations = args.map((arg) => evaluateFormulaCondition(arg, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries));
        if (functionName === "AND") {
            if (evaluations.some((value) => value === false)) {
                return "FALSE";
            }
            if (evaluations.some((value) => value == null)) {
                return null;
            }
            const booleans = evaluations;
            return booleans.every(Boolean) ? "TRUE" : "FALSE";
        }
        if (functionName === "OR") {
            if (evaluations.some((value) => value === true)) {
                return "TRUE";
            }
            if (evaluations.some((value) => value == null)) {
                return null;
            }
            const booleans = evaluations;
            return booleans.some(Boolean) ? "TRUE" : "FALSE";
        }
        return null;
    }
    function tryResolveTextFunction(normalizedFormula, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries) {
        const call = parseWholeFunctionCall(normalizedFormula, ["TEXT"]);
        if (!call)
            return null;
        const args = splitFormulaArguments(call.argsText.trim());
        if (args.length !== 2)
            return null;
        const value = resolveScalarFormulaValue(args[0], currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
        const formatText = resolveScalarFormulaValue(args[1], currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
        if (value == null || formatText == null)
            return null;
        return formatTextFunctionValue(value, formatText);
    }
    function tryResolveLookupFunction(normalizedFormula, currentSheetName, resolveCellValue) {
        var _a;
        const xlookupCall = parseWholeFunctionCall(normalizedFormula, ["XLOOKUP"]);
        if (xlookupCall) {
            const args = splitFormulaArguments(xlookupCall.argsText.trim());
            if (args.length < 3 || args.length > 6)
                return null;
            const lookupValue = resolveScalarFormulaValue(args[0], currentSheetName, resolveCellValue);
            const lookupRange = parseQualifiedRangeReference(args[1], currentSheetName);
            const returnRange = parseQualifiedRangeReference(args[2], currentSheetName);
            if (lookupValue == null || !lookupRange || !returnRange)
                return null;
            const lookupCells = collectRangeCells(lookupRange, resolveCellValue);
            const returnCells = collectRangeCells(returnRange, resolveCellValue);
            if (lookupCells.length === 0 || lookupCells.length !== returnCells.length)
                return null;
            if (args.length >= 5) {
                const matchMode = resolveScalarFormulaValue(args[4], currentSheetName, resolveCellValue);
                if (matchMode == null || !["0", ""].includes(matchMode.trim())) {
                    return null;
                }
            }
            if (args.length >= 6) {
                const searchMode = resolveScalarFormulaValue(args[5], currentSheetName, resolveCellValue);
                if (searchMode == null || !["1", ""].includes(searchMode.trim())) {
                    return null;
                }
            }
            for (let index = 0; index < lookupCells.length; index += 1) {
                const value = lookupCells[index];
                if (value === lookupValue || (!Number.isNaN(Number(value)) && !Number.isNaN(Number(lookupValue)) && Number(value) === Number(lookupValue))) {
                    return (_a = returnCells[index]) !== null && _a !== void 0 ? _a : "";
                }
            }
            if (args.length >= 4) {
                return resolveScalarFormulaValue(args[3], currentSheetName, resolveCellValue);
            }
            return null;
        }
        const matchCall = parseWholeFunctionCall(normalizedFormula, ["MATCH"]);
        if (matchCall) {
            const args = splitFormulaArguments(matchCall.argsText.trim());
            if (args.length < 2 || args.length > 3)
                return null;
            const lookupValue = resolveScalarFormulaValue(args[0], currentSheetName, resolveCellValue);
            const rangeRef = parseQualifiedRangeReference(args[1], currentSheetName);
            if (lookupValue == null || !rangeRef)
                return null;
            if (args.length === 3) {
                const matchType = resolveScalarFormulaValue(args[2], currentSheetName, resolveCellValue);
                if (matchType == null || !["0", ""].includes(matchType.trim())) {
                    return null;
                }
            }
            const cells = collectRangeCells(rangeRef, resolveCellValue);
            if (cells.length === 0)
                return null;
            for (let index = 0; index < cells.length; index += 1) {
                const value = cells[index];
                if (value === lookupValue || (!Number.isNaN(Number(value)) && !Number.isNaN(Number(lookupValue)) && Number(value) === Number(lookupValue))) {
                    return String(index + 1);
                }
            }
            return null;
        }
        const indexCall = parseWholeFunctionCall(normalizedFormula, ["INDEX"]);
        if (indexCall) {
            const args = splitFormulaArguments(indexCall.argsText.trim());
            if (args.length < 2 || args.length > 3)
                return null;
            const rangeRef = parseQualifiedRangeReference(args[0], currentSheetName);
            const rowIndex = Number(resolveScalarFormulaValue(args[1], currentSheetName, resolveCellValue));
            const colIndex = args.length === 3
                ? Number(resolveScalarFormulaValue(args[2], currentSheetName, resolveCellValue))
                : 1;
            if (!rangeRef || Number.isNaN(rowIndex) || Number.isNaN(colIndex) || rowIndex < 1 || colIndex < 1)
                return null;
            const start = parseCellAddress(rangeRef.start);
            const end = parseCellAddress(rangeRef.end);
            if (!start.row || !start.col || !end.row || !end.col)
                return null;
            const startRow = Math.min(start.row, end.row);
            const endRow = Math.max(start.row, end.row);
            const startCol = Math.min(start.col, end.col);
            const endCol = Math.max(start.col, end.col);
            const targetRow = startRow + Math.trunc(rowIndex) - 1;
            const targetCol = startCol + Math.trunc(colIndex) - 1;
            if (targetRow > endRow || targetCol > endCol)
                return null;
            return resolveCellValue(rangeRef.sheetName, `${colToLetters(targetCol)}${targetRow}`);
        }
        const hlookupCall = parseWholeFunctionCall(normalizedFormula, ["HLOOKUP"]);
        if (hlookupCall) {
            const args = splitFormulaArguments(hlookupCall.argsText.trim());
            if (args.length < 3 || args.length > 4)
                return null;
            const lookupValue = resolveScalarFormulaValue(args[0], currentSheetName, resolveCellValue);
            const rangeRef = parseQualifiedRangeReference(args[1], currentSheetName);
            const rowIndex = Number(resolveScalarFormulaValue(args[2], currentSheetName, resolveCellValue));
            if (lookupValue == null || !rangeRef || Number.isNaN(rowIndex) || rowIndex < 1)
                return null;
            if (args.length === 4) {
                const rangeLookup = resolveScalarFormulaValue(args[3], currentSheetName, resolveCellValue);
                if (rangeLookup == null)
                    return null;
                const normalizedLookup = rangeLookup.trim().toUpperCase();
                if (!(normalizedLookup === "FALSE" || normalizedLookup === "0" || normalizedLookup === "")) {
                    return null;
                }
            }
            const start = parseCellAddress(rangeRef.start);
            const end = parseCellAddress(rangeRef.end);
            if (!start.row || !start.col || !end.row || !end.col)
                return null;
            const startRow = Math.min(start.row, end.row);
            const endRow = Math.max(start.row, end.row);
            const startCol = Math.min(start.col, end.col);
            const endCol = Math.max(start.col, end.col);
            const targetRow = startRow + Math.trunc(rowIndex) - 1;
            if (targetRow > endRow)
                return null;
            for (let col = startCol; col <= endCol; col += 1) {
                const keyValue = resolveCellValue(rangeRef.sheetName, `${colToLetters(col)}${startRow}`);
                if (keyValue === "")
                    continue;
                if (keyValue === lookupValue || (!Number.isNaN(Number(keyValue)) && !Number.isNaN(Number(lookupValue)) && Number(keyValue) === Number(lookupValue))) {
                    return resolveCellValue(rangeRef.sheetName, `${colToLetters(col)}${targetRow}`);
                }
            }
            return null;
        }
        const call = parseWholeFunctionCall(normalizedFormula, ["VLOOKUP"]);
        if (!call)
            return null;
        const args = splitFormulaArguments(call.argsText.trim());
        if (args.length < 3 || args.length > 4)
            return null;
        const lookupValue = resolveScalarFormulaValue(args[0], currentSheetName, resolveCellValue);
        const rangeRef = parseQualifiedRangeReference(args[1], currentSheetName);
        const columnIndex = Number(resolveScalarFormulaValue(args[2], currentSheetName, resolveCellValue));
        if (lookupValue == null || !rangeRef || Number.isNaN(columnIndex))
            return null;
        if (columnIndex < 1)
            return null;
        if (args.length === 4) {
            const rangeLookup = resolveScalarFormulaValue(args[3], currentSheetName, resolveCellValue);
            if (rangeLookup == null)
                return null;
            const normalizedLookup = rangeLookup.trim().toUpperCase();
            if (!(normalizedLookup === "FALSE" || normalizedLookup === "0" || normalizedLookup === "")) {
                return null;
            }
        }
        const start = parseCellAddress(rangeRef.start);
        const end = parseCellAddress(rangeRef.end);
        if (!start.row || !start.col || !end.row || !end.col)
            return null;
        const startRow = Math.min(start.row, end.row);
        const endRow = Math.max(start.row, end.row);
        const startCol = Math.min(start.col, end.col);
        const endCol = Math.max(start.col, end.col);
        const targetCol = startCol + Math.trunc(columnIndex) - 1;
        if (targetCol > endCol)
            return null;
        for (let row = startRow; row <= endRow; row += 1) {
            const keyValue = resolveCellValue(rangeRef.sheetName, `${colToLetters(startCol)}${row}`);
            if (keyValue === "")
                continue;
            if (keyValue === lookupValue || (!Number.isNaN(Number(keyValue)) && !Number.isNaN(Number(lookupValue)) && Number(keyValue) === Number(lookupValue))) {
                return resolveCellValue(rangeRef.sheetName, `${colToLetters(targetCol)}${row}`);
            }
        }
        return null;
    }
    function tryResolveDatePartFunction(normalizedFormula, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries) {
        const call = parseWholeFunctionCall(normalizedFormula, ["YEAR", "MONTH", "DAY", "WEEKDAY"]);
        if (!call)
            return null;
        const fnName = call.name;
        const args = splitFormulaArguments(call.argsText.trim());
        if ((fnName === "WEEKDAY" && (args.length < 1 || args.length > 2)) || (fnName !== "WEEKDAY" && args.length !== 1)) {
            return null;
        }
        const value = resolveScalarFormulaValue(args[0], currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
        if (value == null)
            return null;
        const parts = parseDateLikeParts(value);
        if (!parts)
            return null;
        if (fnName === "YEAR")
            return String(Number(parts.yyyy));
        if (fnName === "MONTH")
            return String(Number(parts.mm));
        if (fnName === "DAY")
            return String(Number(parts.dd));
        if (fnName === "WEEKDAY") {
            const returnType = args.length === 2
                ? resolveNumericFormulaArgument(args[1], currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries)
                : 1;
            if (returnType == null)
                return null;
            const weekday = new Date(Date.UTC(Number(parts.yyyy), Number(parts.mm) - 1, Number(parts.dd))).getUTCDay();
            if (Math.trunc(returnType) === 2) {
                return String(weekday === 0 ? 7 : weekday);
            }
            return String(weekday + 1);
        }
        return null;
    }
    function tryResolvePredicateFunction(normalizedFormula, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries) {
        const call = parseWholeFunctionCall(normalizedFormula, ["ISBLANK", "ISNUMBER", "ISTEXT", "ISERROR", "ISNA"]);
        if (!call)
            return null;
        const fnName = call.name;
        const args = splitFormulaArguments(call.argsText.trim());
        if (args.length !== 1)
            return null;
        if (fnName === "ISBLANK") {
            const simpleRef = parseSimpleFormulaReference(`=${args[0].trim()}`, currentSheetName);
            if (simpleRef) {
                const value = resolveCellValue(simpleRef.sheetName, simpleRef.address);
                return value.trim() === "" ? "TRUE" : "FALSE";
            }
            const value = resolveScalarFormulaValue(args[0], currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
            return value == null || value.trim() === "" ? "TRUE" : "FALSE";
        }
        const value = resolveScalarFormulaValue(args[0], currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
        if (fnName === "ISERROR") {
            if (value == null) {
                return "TRUE";
            }
            return /^#(?:[A-Z]+\/[A-Z]+|[A-Z]+[!?]?)/i.test(value.trim()) ? "TRUE" : "FALSE";
        }
        if (fnName === "ISNA") {
            if (/^\s*VLOOKUP\(/i.test(args[0])) {
                return value == null ? "TRUE" : "FALSE";
            }
            if (value == null) {
                return "FALSE";
            }
            return /^#N\/A$/i.test(value.trim()) ? "TRUE" : "FALSE";
        }
        if (value == null) {
            return "FALSE";
        }
        if (fnName === "ISNUMBER") {
            if (value.trim() === "")
                return "FALSE";
            return !Number.isNaN(Number(value)) ? "TRUE" : "FALSE";
        }
        if (fnName === "ISTEXT") {
            const normalized = value.trim().toUpperCase();
            if (normalized === "")
                return "FALSE";
            if (normalized === "TRUE" || normalized === "FALSE")
                return "FALSE";
            return Number.isNaN(Number(value)) ? "TRUE" : "FALSE";
        }
        return null;
    }
    function tryResolveChooseFunction(normalizedFormula, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries) {
        const call = parseWholeFunctionCall(normalizedFormula, ["CHOOSE"]);
        if (!call)
            return null;
        const args = splitFormulaArguments(call.argsText.trim());
        if (args.length < 2)
            return null;
        const indexValue = resolveNumericFormulaArgument(args[0], currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
        if (indexValue == null)
            return null;
        const index = Math.trunc(indexValue);
        if (index < 1 || index >= args.length)
            return null;
        return resolveScalarFormulaValue(args[index], currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
    }
    function tryResolveConcatenationExpression(normalizedFormula, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries) {
        const segments = splitConcatenationExpression(normalizedFormula);
        if (!segments || segments.length < 2)
            return null;
        const values = [];
        for (const segment of segments) {
            const resolved = resolveScalarFormulaValue(segment, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
            if (resolved == null) {
                return null;
            }
            values.push(resolved);
        }
        return values.join("");
    }
    function evaluateFormulaCondition(expression, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries) {
        const logical = tryResolveLogicalFunction(expression.trim(), currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
        if (logical != null) {
            return logical === "TRUE";
        }
        const comparison = tryResolveComparisonExpression(expression, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
        if (comparison != null) {
            return comparison === "TRUE";
        }
        const scalar = resolveScalarFormulaValue(expression, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
        if (scalar == null)
            return null;
        const normalized = scalar.trim().toUpperCase();
        if (normalized === "TRUE")
            return true;
        if (normalized === "FALSE")
            return false;
        const numeric = Number(scalar);
        if (!Number.isNaN(numeric)) {
            return numeric !== 0;
        }
        return scalar.trim() !== "";
    }
    function tryResolveComparisonExpression(normalizedFormula, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries) {
        const comparison = splitComparisonExpression(normalizedFormula);
        if (!comparison)
            return null;
        const left = resolveScalarFormulaValue(comparison.left, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
        const right = resolveScalarFormulaValue(comparison.right, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
        if (left == null || right == null)
            return null;
        const leftNum = Number(left);
        const rightNum = Number(right);
        const numericComparable = !Number.isNaN(leftNum) && !Number.isNaN(rightNum);
        let result = false;
        if (comparison.operator === "=") {
            result = numericComparable ? leftNum === rightNum : left === right;
        }
        else if (comparison.operator === "<>") {
            result = numericComparable ? leftNum !== rightNum : left !== right;
        }
        else if (!numericComparable) {
            return null;
        }
        else if (comparison.operator === ">") {
            result = leftNum > rightNum;
        }
        else if (comparison.operator === "<") {
            result = leftNum < rightNum;
        }
        else if (comparison.operator === ">=") {
            result = leftNum >= rightNum;
        }
        else if (comparison.operator === "<=") {
            result = leftNum <= rightNum;
        }
        return result ? "TRUE" : "FALSE";
    }
    function splitComparisonExpression(expression) {
        const operators = ["<=", ">=", "<>", "=", ">", "<"];
        let depth = 0;
        let inSingleQuote = false;
        let inDoubleQuote = false;
        for (let i = 0; i < expression.length; i += 1) {
            const ch = expression[i];
            if (ch === "'" && !inDoubleQuote) {
                inSingleQuote = !inSingleQuote;
                continue;
            }
            if (ch === "\"" && !inSingleQuote) {
                inDoubleQuote = !inDoubleQuote;
                continue;
            }
            if (inSingleQuote || inDoubleQuote)
                continue;
            if (ch === "(") {
                depth += 1;
                continue;
            }
            if (ch === ")") {
                depth = Math.max(0, depth - 1);
                continue;
            }
            if (depth !== 0)
                continue;
            for (const operator of operators) {
                if (expression.slice(i, i + operator.length) === operator) {
                    return {
                        left: expression.slice(0, i).trim(),
                        operator,
                        right: expression.slice(i + operator.length).trim()
                    };
                }
            }
        }
        return null;
    }
    function findTopLevelOperatorIndex(expression, operator) {
        const target = String(operator || "");
        if (!target)
            return -1;
        let depth = 0;
        let inSingleQuote = false;
        let inDoubleQuote = false;
        for (let i = 0; i <= expression.length - target.length; i += 1) {
            const ch = expression[i];
            if (ch === "'" && !inDoubleQuote) {
                inSingleQuote = !inSingleQuote;
                continue;
            }
            if (ch === "\"" && !inSingleQuote) {
                inDoubleQuote = !inDoubleQuote;
                continue;
            }
            if (inSingleQuote || inDoubleQuote)
                continue;
            if (ch === "(") {
                depth += 1;
                continue;
            }
            if (ch === ")") {
                depth = Math.max(0, depth - 1);
                continue;
            }
            if (depth === 0 && expression.slice(i, i + target.length) === target) {
                return i;
            }
        }
        return -1;
    }
    function splitConcatenationExpression(expression) {
        const parts = [];
        let start = 0;
        let depth = 0;
        let inSingleQuote = false;
        let inDoubleQuote = false;
        for (let i = 0; i < expression.length; i += 1) {
            const ch = expression[i];
            if (ch === "'" && !inDoubleQuote) {
                inSingleQuote = !inSingleQuote;
                continue;
            }
            if (ch === "\"" && !inSingleQuote) {
                inDoubleQuote = !inDoubleQuote;
                continue;
            }
            if (inSingleQuote || inDoubleQuote)
                continue;
            if (ch === "(") {
                depth += 1;
                continue;
            }
            if (ch === ")") {
                depth = Math.max(0, depth - 1);
                continue;
            }
            if (depth === 0 && ch === "&") {
                parts.push(expression.slice(start, i).trim());
                start = i + 1;
            }
        }
        if (parts.length === 0) {
            return null;
        }
        parts.push(expression.slice(start).trim());
        return parts.every(Boolean) ? parts : null;
    }
    function parseWholeFunctionCall(expression, allowedNames) {
        const trimmed = String(expression || "").trim();
        const nameMatch = trimmed.match(/^([A-Z][A-Z0-9]*)\(/i);
        if (!nameMatch)
            return null;
        const name = nameMatch[1].toUpperCase();
        if (!allowedNames.includes(name))
            return null;
        let depth = 0;
        let inSingleQuote = false;
        let inDoubleQuote = false;
        for (let i = name.length; i < trimmed.length; i += 1) {
            const ch = trimmed[i];
            if (ch === "'" && !inDoubleQuote) {
                inSingleQuote = !inSingleQuote;
                continue;
            }
            if (ch === "\"" && !inSingleQuote) {
                inDoubleQuote = !inDoubleQuote;
                continue;
            }
            if (inSingleQuote || inDoubleQuote)
                continue;
            if (ch === "(") {
                depth += 1;
                continue;
            }
            if (ch !== ")")
                continue;
            depth -= 1;
            if (depth > 0)
                continue;
            if (depth < 0 || i !== trimmed.length - 1)
                return null;
            return {
                name,
                argsText: trimmed.slice(name.length + 1, i)
            };
        }
        return null;
    }
    function replaceNumericDefinedNames(expression, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries) {
        let result = "";
        let i = 0;
        let inSingleQuote = false;
        let inDoubleQuote = false;
        while (i < expression.length) {
            const ch = expression[i];
            if (ch === "'" && !inDoubleQuote) {
                inSingleQuote = !inSingleQuote;
                result += ch;
                i += 1;
                continue;
            }
            if (ch === "\"" && !inSingleQuote) {
                inDoubleQuote = !inDoubleQuote;
                result += ch;
                i += 1;
                continue;
            }
            if (inSingleQuote || inDoubleQuote) {
                result += ch;
                i += 1;
                continue;
            }
            if (!/[\p{L}_]/u.test(ch)) {
                result += ch;
                i += 1;
                continue;
            }
            const start = i;
            i += 1;
            while (i < expression.length && /[\p{L}\p{N}_.]/u.test(expression[i])) {
                i += 1;
            }
            const token = expression.slice(start, i);
            const nextChar = expression[i] || "";
            if (nextChar === "(") {
                result += token;
                continue;
            }
            const scalar = (resolveDefinedNameScalarValue === null || resolveDefinedNameScalarValue === void 0 ? void 0 : resolveDefinedNameScalarValue(currentSheetName, token)) || null;
            if (scalar != null) {
                const numeric = Number(scalar);
                if (!Number.isNaN(numeric)) {
                    result += String(numeric);
                    continue;
                }
            }
            result += token;
        }
        return result;
    }
    function replaceEmbeddedNumericFunctions(expression, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries) {
        let current = expression;
        let changed = true;
        while (changed) {
            changed = false;
            current = current.replace(/[A-Z][A-Z0-9]*\([^()]*\)/gi, (segment) => {
                var _a, _b, _c, _d;
                const resolved = (_d = (_c = (_b = (_a = tryResolveNumericFunction(segment, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries)) !== null && _a !== void 0 ? _a : tryResolveDatePartFunction(segment, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries)) !== null && _b !== void 0 ? _b : tryResolveAggregateFunction(segment, currentSheetName, resolveRangeValues, resolveRangeEntries)) !== null && _c !== void 0 ? _c : tryResolveConditionalAggregateFunction(segment, currentSheetName, resolveCellValue)) !== null && _d !== void 0 ? _d : tryResolveLookupFunction(segment, currentSheetName, resolveCellValue);
                if (resolved == null) {
                    return segment;
                }
                const numericValue = Number(resolved);
                if (Number.isNaN(numericValue)) {
                    return segment;
                }
                changed = true;
                return String(numericValue);
            });
        }
        return current;
    }
    function resolveScalarFormulaValue(expression, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries) {
        const trimmed = String(expression || "").trim();
        if (!trimmed)
            return null;
        const quotedString = trimmed.match(/^"(.*)"$/);
        if (quotedString) {
            return quotedString[1].replace(/""/g, "\"");
        }
        const numeric = Number(trimmed);
        if (!Number.isNaN(numeric)) {
            return String(numeric);
        }
        const simpleRef = parseSimpleFormulaReference(`=${trimmed}`, currentSheetName);
        if (simpleRef) {
            return resolveCellValue(simpleRef.sheetName, simpleRef.address);
        }
        const scopedDefinedNameRef = parseSheetScopedDefinedNameReference(trimmed, currentSheetName);
        if (scopedDefinedNameRef) {
            const scopedValue = (resolveDefinedNameScalarValue === null || resolveDefinedNameScalarValue === void 0 ? void 0 : resolveDefinedNameScalarValue(scopedDefinedNameRef.sheetName, scopedDefinedNameRef.name)) || null;
            if (scopedValue != null) {
                return scopedValue;
            }
        }
        const definedNameValue = (resolveDefinedNameScalarValue === null || resolveDefinedNameScalarValue === void 0 ? void 0 : resolveDefinedNameScalarValue(currentSheetName, trimmed)) || null;
        if (definedNameValue != null) {
            return definedNameValue;
        }
        if (/^(TRUE|FALSE)$/i.test(trimmed)) {
            return trimmed.toUpperCase();
        }
        return tryResolveFormulaExpression(`=${trimmed}`, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
    }
    function tryResolveAggregateFunction(normalizedFormula, currentSheetName, resolveRangeValues, resolveRangeEntries) {
        if (!resolveRangeValues || !resolveRangeEntries)
            return null;
        const call = parseWholeFunctionCall(normalizedFormula, ["SUM", "AVERAGE", "MIN", "MAX", "COUNT", "COUNTA"]);
        if (!call)
            return null;
        const fnName = call.name;
        const argText = call.argsText.trim();
        const args = splitFormulaArguments(argText);
        if (args.length === 0) {
            return null;
        }
        const resolvedArgs = args.map((arg) => resolveAggregateArgument(arg, currentSheetName, resolveRangeValues, resolveRangeEntries));
        if (resolvedArgs.some((entry) => entry == null)) {
            return null;
        }
        const values = resolvedArgs.flatMap((entry) => (entry === null || entry === void 0 ? void 0 : entry.numericValues) || []);
        const valueCount = resolvedArgs.reduce((sum, entry) => sum + ((entry === null || entry === void 0 ? void 0 : entry.valueCount) || 0), 0);
        if ((fnName !== "COUNTA" && values.length === 0) || valueCount === 0) {
            return null;
        }
        if (fnName === "SUM") {
            return String(values.reduce((sum, value) => sum + value, 0));
        }
        if (fnName === "AVERAGE") {
            return String(values.reduce((sum, value) => sum + value, 0) / values.length);
        }
        if (fnName === "MIN") {
            return String(Math.min(...values));
        }
        if (fnName === "MAX") {
            return String(Math.max(...values));
        }
        if (fnName === "COUNT") {
            return String(values.length);
        }
        if (fnName === "COUNTA") {
            return String(valueCount);
        }
        return null;
    }
    function tryResolveConditionalAggregateFunction(normalizedFormula, currentSheetName, resolveCellValue) {
        const averageifsCall = parseWholeFunctionCall(normalizedFormula, ["AVERAGEIFS"]);
        if (averageifsCall) {
            const args = splitFormulaArguments(averageifsCall.argsText.trim());
            if (args.length < 3 || args.length % 2 === 0)
                return null;
            const averageRange = parseQualifiedRangeReference(args[0], currentSheetName);
            if (!averageRange)
                return null;
            const averageCells = collectRangeCells(averageRange, resolveCellValue);
            if (averageCells.length === 0)
                return null;
            const rangeCriteriaPairs = [];
            for (let index = 1; index < args.length; index += 2) {
                const rangeRef = parseQualifiedRangeReference(args[index], currentSheetName);
                const criteria = resolveScalarFormulaValue(args[index + 1], currentSheetName, resolveCellValue);
                if (!rangeRef || criteria == null)
                    return null;
                const cells = collectRangeCells(rangeRef, resolveCellValue);
                if (cells.length !== averageCells.length)
                    return null;
                rangeCriteriaPairs.push({ cells, criteria });
            }
            let sum = 0;
            let count = 0;
            for (let i = 0; i < averageCells.length; i += 1) {
                if (!rangeCriteriaPairs.every((entry) => matchesCountIfCriteria(entry.cells[i], entry.criteria))) {
                    continue;
                }
                const numeric = Number(averageCells[i]);
                if (!Number.isNaN(numeric)) {
                    sum += numeric;
                    count += 1;
                }
            }
            return count > 0 ? String(sum / count) : null;
        }
        const sumifsCall = parseWholeFunctionCall(normalizedFormula, ["SUMIFS"]);
        if (sumifsCall) {
            const args = splitFormulaArguments(sumifsCall.argsText.trim());
            if (args.length < 3 || args.length % 2 === 0)
                return null;
            const sumRange = parseQualifiedRangeReference(args[0], currentSheetName);
            if (!sumRange)
                return null;
            const sumCells = collectRangeCells(sumRange, resolveCellValue);
            if (sumCells.length === 0)
                return null;
            const rangeCriteriaPairs = [];
            for (let index = 1; index < args.length; index += 2) {
                const rangeRef = parseQualifiedRangeReference(args[index], currentSheetName);
                const criteria = resolveScalarFormulaValue(args[index + 1], currentSheetName, resolveCellValue);
                if (!rangeRef || criteria == null)
                    return null;
                const cells = collectRangeCells(rangeRef, resolveCellValue);
                if (cells.length !== sumCells.length)
                    return null;
                rangeCriteriaPairs.push({ cells, criteria });
            }
            let sum = 0;
            for (let i = 0; i < sumCells.length; i += 1) {
                if (!rangeCriteriaPairs.every((entry) => matchesCountIfCriteria(entry.cells[i], entry.criteria))) {
                    continue;
                }
                const numeric = Number(sumCells[i]);
                if (!Number.isNaN(numeric)) {
                    sum += numeric;
                }
            }
            return String(sum);
        }
        const countifsCall = parseWholeFunctionCall(normalizedFormula, ["COUNTIFS"]);
        if (countifsCall) {
            const args = splitFormulaArguments(countifsCall.argsText.trim());
            if (args.length < 2 || args.length % 2 !== 0)
                return null;
            const rangeCriteriaPairs = [];
            for (let index = 0; index < args.length; index += 2) {
                const rangeRef = parseQualifiedRangeReference(args[index], currentSheetName);
                const criteria = resolveScalarFormulaValue(args[index + 1], currentSheetName, resolveCellValue);
                if (!rangeRef || criteria == null)
                    return null;
                const cells = collectRangeCells(rangeRef, resolveCellValue);
                if (cells.length === 0)
                    return null;
                rangeCriteriaPairs.push({ cells, criteria });
            }
            const length = rangeCriteriaPairs[0].cells.length;
            if (rangeCriteriaPairs.some((entry) => entry.cells.length !== length))
                return null;
            let count = 0;
            for (let i = 0; i < length; i += 1) {
                if (rangeCriteriaPairs.every((entry) => matchesCountIfCriteria(entry.cells[i], entry.criteria))) {
                    count += 1;
                }
            }
            return String(count);
        }
        const call = parseWholeFunctionCall(normalizedFormula, ["COUNTIF", "SUMIF", "AVERAGEIF"]);
        if (!call)
            return null;
        const fnName = call.name;
        const args = splitFormulaArguments(call.argsText.trim());
        if ((fnName === "COUNTIF" && args.length !== 2) || ((fnName === "SUMIF" || fnName === "AVERAGEIF") && args.length !== 2 && args.length !== 3)) {
            return null;
        }
        const criteriaRange = parseQualifiedRangeReference(args[0], currentSheetName);
        if (!criteriaRange)
            return null;
        const criteria = resolveScalarFormulaValue(args[1], currentSheetName, resolveCellValue);
        if (criteria == null)
            return null;
        const criteriaCells = collectRangeCells(criteriaRange, resolveCellValue);
        if (criteriaCells.length === 0)
            return null;
        const sumRange = fnName === "COUNTIF"
            ? criteriaRange
            : (fnName === "SUMIF" || fnName === "AVERAGEIF")
                ? parseQualifiedRangeReference(args[2] || args[0], currentSheetName)
                : criteriaRange;
        if (!sumRange)
            return null;
        const sumCells = collectRangeCells(sumRange, resolveCellValue);
        if (sumCells.length !== criteriaCells.length)
            return null;
        let count = 0;
        let sum = 0;
        for (let i = 0; i < criteriaCells.length; i += 1) {
            if (!matchesCountIfCriteria(criteriaCells[i], criteria))
                continue;
            count += 1;
            const numeric = Number(sumCells[i]);
            if (!Number.isNaN(numeric)) {
                sum += numeric;
            }
        }
        if (fnName === "COUNTIF") {
            return String(count);
        }
        if (fnName === "SUMIF") {
            return String(sum);
        }
        return count > 0 ? String(sum / count) : null;
    }
    function tryResolveNumericFunction(normalizedFormula, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries) {
        const call = parseWholeFunctionCall(normalizedFormula, ["ROUND", "ROUNDUP", "ROUNDDOWN", "INT", "DATE", "VALUE", "DATEVALUE", "ROW", "COLUMN", "EOMONTH"]);
        if (!call)
            return null;
        const fnName = call.name;
        const args = splitFormulaArguments(call.argsText.trim());
        if (fnName === "ROW" || fnName === "COLUMN") {
            if (args.length !== 1)
                return null;
            const rangeRef = parseQualifiedRangeReference(args[0], currentSheetName);
            if (rangeRef) {
                const start = parseCellAddress(rangeRef.start);
                if (!start.row || !start.col)
                    return null;
                return String(fnName === "ROW" ? start.row : start.col);
            }
            const simpleRef = parseSimpleFormulaReference(`=${args[0]}`, currentSheetName);
            if (!simpleRef)
                return null;
            const parsed = parseCellAddress(simpleRef.address);
            if (!parsed.row || !parsed.col)
                return null;
            return String(fnName === "ROW" ? parsed.row : parsed.col);
        }
        if (fnName === "VALUE" || fnName === "DATEVALUE") {
            if (args.length !== 1)
                return null;
            const source = resolveScalarFormulaValue(args[0], currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
            if (source == null)
                return null;
            const parsed = parseValueFunctionText(source);
            return parsed == null ? null : String(parsed);
        }
        if (fnName === "DATE") {
            if (args.length !== 3)
                return null;
            const year = resolveNumericFormulaArgument(args[0], currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
            const month = resolveNumericFormulaArgument(args[1], currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
            const day = resolveNumericFormulaArgument(args[2], currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
            if (year == null || month == null || day == null)
                return null;
            const serial = datePartsToExcelSerial(Math.trunc(year), Math.trunc(month), Math.trunc(day));
            return serial == null ? null : String(serial);
        }
        if (fnName === "EOMONTH") {
            if (args.length !== 2)
                return null;
            const startValue = resolveScalarFormulaValue(args[0], currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
            const monthOffset = resolveNumericFormulaArgument(args[1], currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
            if (startValue == null || monthOffset == null)
                return null;
            const parts = parseDateLikeParts(startValue);
            if (!parts)
                return null;
            const baseYear = Number(parts.yyyy);
            const baseMonthIndex = Number(parts.mm) - 1 + Math.trunc(monthOffset);
            const targetYear = baseYear + Math.floor(baseMonthIndex / 12);
            const targetMonth = ((baseMonthIndex % 12) + 12) % 12 + 1;
            const serial = datePartsToExcelSerial(targetYear, targetMonth + 1, 0);
            return serial == null ? null : String(serial);
        }
        if (fnName === "INT") {
            if (args.length !== 1)
                return null;
            const value = resolveNumericFormulaArgument(args[0], currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
            if (value == null)
                return null;
            return String(Math.floor(value));
        }
        if (args.length !== 2)
            return null;
        const value = resolveNumericFormulaArgument(args[0], currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
        const digits = resolveNumericFormulaArgument(args[1], currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
        if (value == null || digits == null)
            return null;
        const roundedDigits = Math.trunc(digits);
        const factor = 10 ** roundedDigits;
        if (!Number.isFinite(factor) || factor === 0)
            return null;
        if (fnName === "ROUND") {
            return String(Math.round(value * factor) / factor);
        }
        if (fnName === "ROUNDUP") {
            const scaled = value * factor;
            const rounded = scaled >= 0 ? Math.ceil(scaled) : Math.floor(scaled);
            return String(rounded / factor);
        }
        if (fnName === "ROUNDDOWN") {
            const scaled = value * factor;
            const rounded = scaled >= 0 ? Math.floor(scaled) : Math.ceil(scaled);
            return String(rounded / factor);
        }
        return null;
    }
    function tryResolveStringFunction(normalizedFormula, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries) {
        const call = parseWholeFunctionCall(normalizedFormula, ["LEFT", "RIGHT", "MID", "LEN", "TRIM", "SUBSTITUTE", "REPLACE", "REPT"]);
        if (!call)
            return null;
        const fnName = call.name;
        const args = splitFormulaArguments(call.argsText.trim());
        if (fnName === "LEN" || fnName === "TRIM") {
            if (args.length !== 1)
                return null;
            const source = resolveScalarFormulaValue(args[0], currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
            if (source == null)
                return null;
            if (fnName === "LEN") {
                return String(source.length);
            }
            return source.trim().replace(/\s+/g, " ");
        }
        if (fnName === "LEFT" || fnName === "RIGHT") {
            if (args.length < 1 || args.length > 2)
                return null;
            const source = resolveScalarFormulaValue(args[0], currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
            if (source == null)
                return null;
            const count = args.length === 2
                ? resolveNumericFormulaArgument(args[1], currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries)
                : 1;
            if (count == null)
                return null;
            const length = Math.max(0, Math.trunc(count));
            return fnName === "LEFT" ? source.slice(0, length) : source.slice(Math.max(0, source.length - length));
        }
        if (fnName === "MID") {
            if (args.length !== 3)
                return null;
            const source = resolveScalarFormulaValue(args[0], currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
            const start = resolveNumericFormulaArgument(args[1], currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
            const count = resolveNumericFormulaArgument(args[2], currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
            if (source == null || start == null || count == null)
                return null;
            const startIndex = Math.max(0, Math.trunc(start) - 1);
            const length = Math.max(0, Math.trunc(count));
            return source.slice(startIndex, startIndex + length);
        }
        if (fnName === "SUBSTITUTE") {
            if (args.length < 3 || args.length > 4)
                return null;
            const source = resolveScalarFormulaValue(args[0], currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
            const oldText = resolveScalarFormulaValue(args[1], currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
            const newText = resolveScalarFormulaValue(args[2], currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
            if (source == null || oldText == null || newText == null || oldText === "")
                return null;
            if (args.length === 3) {
                return source.split(oldText).join(newText);
            }
            const instanceNum = resolveNumericFormulaArgument(args[3], currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
            if (instanceNum == null)
                return null;
            const targetIndex = Math.trunc(instanceNum);
            if (targetIndex <= 0)
                return source;
            let occurrence = 0;
            let cursor = 0;
            let result = "";
            while (cursor < source.length) {
                const found = source.indexOf(oldText, cursor);
                if (found === -1) {
                    result += source.slice(cursor);
                    break;
                }
                occurrence += 1;
                result += source.slice(cursor, found);
                if (occurrence === targetIndex) {
                    result += newText;
                    result += source.slice(found + oldText.length);
                    return result;
                }
                result += oldText;
                cursor = found + oldText.length;
            }
            return result || source;
        }
        if (fnName === "REPLACE") {
            if (args.length !== 4)
                return null;
            const source = resolveScalarFormulaValue(args[0], currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
            const start = resolveNumericFormulaArgument(args[1], currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
            const count = resolveNumericFormulaArgument(args[2], currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
            const replacement = resolveScalarFormulaValue(args[3], currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
            if (source == null || start == null || count == null || replacement == null)
                return null;
            const startIndex = Math.max(0, Math.trunc(start) - 1);
            const length = Math.max(0, Math.trunc(count));
            return source.slice(0, startIndex) + replacement + source.slice(startIndex + length);
        }
        if (fnName === "REPT") {
            if (args.length !== 2)
                return null;
            const source = resolveScalarFormulaValue(args[0], currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
            const countValue = resolveScalarFormulaValue(args[1], currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
            if (source == null)
                return null;
            const normalizedCount = countValue == null
                ? (() => {
                    const evaluatedCondition = evaluateFormulaCondition(args[1], currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
                    if (evaluatedCondition == null) {
                        return null;
                    }
                    return evaluatedCondition ? "TRUE" : "FALSE";
                })()
                : countValue.trim().toUpperCase();
            if (normalizedCount == null)
                return null;
            const count = normalizedCount === "TRUE"
                ? 1
                : normalizedCount === "FALSE"
                    ? 0
                    : Number(countValue);
            if (!Number.isFinite(count))
                return null;
            return source.repeat(Math.max(0, Math.trunc(count)));
        }
        return null;
    }
    function splitFormulaArguments(argText) {
        const args = [];
        let current = "";
        let depth = 0;
        let inSingleQuote = false;
        let inDoubleQuote = false;
        for (let i = 0; i < argText.length; i += 1) {
            const ch = argText[i];
            if (ch === "'" && !inDoubleQuote) {
                inSingleQuote = !inSingleQuote;
                current += ch;
                continue;
            }
            if (ch === "\"" && !inSingleQuote) {
                inDoubleQuote = !inDoubleQuote;
                current += ch;
                continue;
            }
            if (!inSingleQuote && !inDoubleQuote) {
                if (ch === "(") {
                    depth += 1;
                }
                else if (ch === ")") {
                    depth = Math.max(0, depth - 1);
                }
                else if (ch === "," && depth === 0) {
                    args.push(current.trim());
                    current = "";
                    continue;
                }
            }
            current += ch;
        }
        if (current.trim()) {
            args.push(current.trim());
        }
        return args;
    }
    function resolveAggregateArgument(argText, currentSheetName, resolveRangeValues, resolveRangeEntries) {
        const rangeRef = parseQualifiedRangeReference(argText, currentSheetName);
        if (rangeRef) {
            const sheetName = rangeRef.sheetName;
            const rangeText = `${rangeRef.start}:${rangeRef.end}`;
            const rangeEntries = resolveRangeEntries(sheetName, rangeText);
            return {
                numericValues: rangeEntries.numericValues,
                valueCount: rangeEntries.rawValues.filter((value) => String(value || "").trim() !== "").length
            };
        }
        const numericLiteral = Number(argText);
        if (!Number.isNaN(numericLiteral)) {
            return {
                numericValues: [numericLiteral],
                valueCount: 1
            };
        }
        const cellRef = parseSimpleFormulaReference(`=${argText}`, currentSheetName);
        if (cellRef) {
            const values = resolveRangeValues(cellRef.sheetName, `${cellRef.address}:${cellRef.address}`);
            const entryCount = resolveRangeEntries(cellRef.sheetName, `${cellRef.address}:${cellRef.address}`).rawValues
                .filter((value) => String(value || "").trim() !== "").length;
            if (values.length > 0) {
                return {
                    numericValues: values,
                    valueCount: entryCount
                };
            }
            return {
                numericValues: [],
                valueCount: entryCount
            };
        }
        return null;
    }
    function resolveNumericFormulaArgument(expression, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries) {
        const scalar = resolveScalarFormulaValue(expression, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
        if (scalar == null)
            return null;
        const numeric = Number(scalar);
        return Number.isNaN(numeric) ? null : numeric;
    }
    function collectRangeCells(rangeRef, resolveCellValue) {
        const start = parseCellAddress(rangeRef.start);
        const end = parseCellAddress(rangeRef.end);
        if (!start.row || !start.col || !end.row || !end.col)
            return [];
        const startRow = Math.min(start.row, end.row);
        const endRow = Math.max(start.row, end.row);
        const startCol = Math.min(start.col, end.col);
        const endCol = Math.max(start.col, end.col);
        const values = [];
        for (let row = startRow; row <= endRow; row += 1) {
            for (let col = startCol; col <= endCol; col += 1) {
                values.push(resolveCellValue(rangeRef.sheetName, `${colToLetters(col)}${row}`));
            }
        }
        return values;
    }
    function matchesCountIfCriteria(value, criteria) {
        const trimmedCriteria = String(criteria || "").trim();
        const operatorMatch = trimmedCriteria.match(/^(<=|>=|<>|=|<|>)(.*)$/);
        const operator = operatorMatch ? operatorMatch[1] : "=";
        const operandText = operatorMatch ? operatorMatch[2].trim() : trimmedCriteria;
        const leftNum = Number(value);
        const rightNum = Number(operandText);
        const numericComparable = !Number.isNaN(leftNum) && !Number.isNaN(rightNum);
        if (operator === "=") {
            return numericComparable ? leftNum === rightNum : value === operandText;
        }
        if (operator === "<>") {
            return numericComparable ? leftNum !== rightNum : value !== operandText;
        }
        if (!numericComparable) {
            return false;
        }
        if (operator === ">")
            return leftNum > rightNum;
        if (operator === "<")
            return leftNum < rightNum;
        if (operator === ">=")
            return leftNum >= rightNum;
        if (operator === "<=")
            return leftNum <= rightNum;
        return false;
    }
    function parseQualifiedRangeReference(argText, currentSheetName) {
        const qualifiedRangeMatch = argText.match(/^(?:'((?:[^']|'')+)'|([^'=][^!]*))!(\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+)$/i);
        const localRangeMatch = argText.match(/^(\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+)$/i);
        if (!qualifiedRangeMatch && !localRangeMatch) {
            const scopedDefinedName = parseSheetScopedDefinedNameReference(String(argText || "").trim(), currentSheetName);
            if (scopedDefinedName) {
                const scopedRange = (resolveDefinedNameRangeRef === null || resolveDefinedNameRangeRef === void 0 ? void 0 : resolveDefinedNameRangeRef(scopedDefinedName.sheetName, scopedDefinedName.name)) || null;
                if (scopedRange) {
                    return scopedRange;
                }
            }
            const structuredRange = (resolveStructuredRangeRef === null || resolveStructuredRangeRef === void 0 ? void 0 : resolveStructuredRangeRef(currentSheetName, String(argText || "").trim())) || null;
            if (structuredRange) {
                return structuredRange;
            }
            const definedRange = (resolveDefinedNameRangeRef === null || resolveDefinedNameRangeRef === void 0 ? void 0 : resolveDefinedNameRangeRef(currentSheetName, String(argText || "").trim())) || null;
            if (definedRange) {
                return definedRange;
            }
            return null;
        }
        const sheetName = qualifiedRangeMatch
            ? normalizeFormulaSheetName(qualifiedRangeMatch[1] || qualifiedRangeMatch[2] || currentSheetName)
            : currentSheetName;
        const rangeText = String(qualifiedRangeMatch ? qualifiedRangeMatch[3] : (localRangeMatch === null || localRangeMatch === void 0 ? void 0 : localRangeMatch[1]) || "");
        const range = parseRangeAddress(rangeText);
        if (!range)
            return null;
        return {
            sheetName,
            start: range.start,
            end: range.end
        };
    }
    function evaluateArithmeticExpression(expression) {
        const tokens = tokenizeArithmeticExpression(expression);
        let index = 0;
        function parseExpression() {
            let value = parseTerm();
            while (tokens[index] === "+" || tokens[index] === "-") {
                const operator = tokens[index];
                index += 1;
                const right = parseTerm();
                value = operator === "+" ? value + right : value - right;
            }
            return value;
        }
        function parseTerm() {
            let value = parseFactor();
            while (tokens[index] === "*" || tokens[index] === "/") {
                const operator = tokens[index];
                index += 1;
                const right = parseFactor();
                value = operator === "*" ? value * right : value / right;
            }
            return value;
        }
        function parseFactor() {
            const token = tokens[index];
            if (token === "+") {
                index += 1;
                return parseFactor();
            }
            if (token === "-") {
                index += 1;
                return -parseFactor();
            }
            if (token === "(") {
                index += 1;
                const value = parseExpression();
                if (tokens[index] !== ")") {
                    throw new Error("Unbalanced parentheses");
                }
                index += 1;
                return value;
            }
            if (token == null) {
                throw new Error("Unexpected end of expression");
            }
            index += 1;
            const numericValue = Number(token);
            if (Number.isNaN(numericValue)) {
                throw new Error("Invalid token");
            }
            return numericValue;
        }
        const result = parseExpression();
        if (index !== tokens.length) {
            throw new Error("Unexpected trailing tokens");
        }
        return result;
    }
    function tokenizeArithmeticExpression(expression) {
        const tokens = [];
        let index = 0;
        while (index < expression.length) {
            const ch = expression[index];
            if (/\s/.test(ch)) {
                index += 1;
                continue;
            }
            if (/[+\-*/()]/.test(ch)) {
                tokens.push(ch);
                index += 1;
                continue;
            }
            const numberMatch = expression.slice(index).match(/^\d+(?:\.\d+)?/);
            if (!numberMatch) {
                throw new Error("Invalid arithmetic expression");
            }
            tokens.push(numberMatch[0]);
            index += numberMatch[0].length;
        }
        return tokens;
    }
    function resolveSimpleFormulaReferences(workbook) {
        var _a, _b;
        const resolver = buildFormulaResolver(workbook);
        resolveDefinedNameScalarValue = resolver.resolveDefinedNameValue;
        resolveDefinedNameRangeRef = resolver.resolveDefinedNameRange;
        resolveStructuredRangeRef = resolver.resolveStructuredRange;
        try {
            for (let pass = 0; pass < 8; pass += 1) {
                let resolvedInPass = 0;
                for (const sheet of workbook.sheets) {
                    for (const cell of sheet.cells) {
                        if (!cell.formulaText)
                            continue;
                        if (cell.resolutionStatus === "unsupported_external")
                            continue;
                        if (cell.resolutionStatus === "resolved")
                            continue;
                        const reference = parseSimpleFormulaReference(cell.formulaText, sheet.name);
                        if (reference) {
                            const targetValue = String(resolver.resolveCellValue(reference.sheetName, reference.address) || "").trim();
                            if (targetValue) {
                                applyResolvedFormulaValue(cell, targetValue, "legacy_resolver");
                                resolvedInPass += 1;
                                continue;
                            }
                        }
                        let evaluated = null;
                        let evaluatedSource = null;
                        try {
                            const result = tryResolveFormulaExpressionDetailed(cell.formulaText, sheet.name, resolver.resolveCellValue, resolver.resolveRangeValues, resolver.resolveRangeEntries, cell.address);
                            evaluated = (_a = result === null || result === void 0 ? void 0 : result.value) !== null && _a !== void 0 ? _a : null;
                            evaluatedSource = (_b = result === null || result === void 0 ? void 0 : result.source) !== null && _b !== void 0 ? _b : null;
                        }
                        catch (error) {
                            if (!(error instanceof Error) || error.message !== "__FORMULA_UNRESOLVED__") {
                                throw error;
                            }
                        }
                        if (evaluated != null) {
                            applyResolvedFormulaValue(cell, evaluated, evaluatedSource || "legacy_resolver");
                            resolvedInPass += 1;
                        }
                    }
                }
                if (resolvedInPass === 0) {
                    break;
                }
            }
        }
        finally {
            resolveDefinedNameScalarValue = null;
            resolveDefinedNameRangeRef = null;
            resolveStructuredRangeRef = null;
        }
    }
    function parseWorksheet(files, sheetName, sheetPath, sheetIndex, sharedStrings, cellStyles) {
        const bytes = files.get(sheetPath);
        if (!bytes) {
            throw new Error(`Sheet XML not found: ${sheetPath}`);
        }
        const doc = xmlToDocument(decodeXmlText(bytes));
        const sharedFormulaMap = new Map();
        const cells = Array.from(doc.getElementsByTagName("c")).map((cellElement) => {
            const address = cellElement.getAttribute("r") || "";
            const position = parseCellAddress(address);
            const styleIndex = Number(cellElement.getAttribute("s") || 0);
            const cellStyle = cellStyles[styleIndex] || {
                borders: EMPTY_BORDERS,
                numFmtId: 0,
                formatCode: "General"
            };
            let formulaOverride = "";
            const formulaElement = cellElement.getElementsByTagName("f")[0] || null;
            const formulaType = (formulaElement === null || formulaElement === void 0 ? void 0 : formulaElement.getAttribute("t")) || "";
            const spillRef = (formulaElement === null || formulaElement === void 0 ? void 0 : formulaElement.getAttribute("ref")) || "";
            const sharedIndex = (formulaElement === null || formulaElement === void 0 ? void 0 : formulaElement.getAttribute("si")) || "";
            const formulaText = getTextContent(formulaElement);
            if (formulaType === "shared" && sharedIndex) {
                if (formulaText) {
                    const normalizedFormula = formulaText.startsWith("=") ? formulaText : `=${formulaText}`;
                    sharedFormulaMap.set(sharedIndex, { address, formulaText: normalizedFormula });
                    formulaOverride = normalizedFormula;
                }
                else {
                    const sharedBase = sharedFormulaMap.get(sharedIndex);
                    if (sharedBase) {
                        formulaOverride = translateSharedFormula(sharedBase.formulaText, sharedBase.address, address);
                    }
                }
            }
            const output = extractCellOutputValue(cellElement, sharedStrings, cellStyle, formulaOverride);
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
                formulaType,
                spillRef
            };
        });
        const merges = Array.from(doc.getElementsByTagName("mergeCell")).map((mergeElement) => parseRangeRef(mergeElement.getAttribute("ref") || ""));
        const tables = worksheetTablesHelper.parseWorksheetTables(files, doc, sheetName, sheetPath);
        const images = parseDrawingImages(files, sheetName, sheetPath);
        const charts = parseDrawingCharts(files, sheetName, sheetPath);
        const shapes = parseDrawingShapes(files, sheetName, sheetPath);
        let maxRow = 0;
        let maxCol = 0;
        for (const cell of cells) {
            if (cell.row > maxRow)
                maxRow = cell.row;
            if (cell.col > maxCol)
                maxCol = cell.col;
        }
        for (const merge of merges) {
            if (merge.endRow > maxRow)
                maxRow = merge.endRow;
            if (merge.endCol > maxCol)
                maxCol = merge.endCol;
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
    async function parseWorkbook(arrayBuffer, workbookName = "workbook.xlsx") {
        const files = await zipIoHelper.unzipEntries(arrayBuffer);
        const workbookBytes = files.get("xl/workbook.xml");
        if (!workbookBytes) {
            throw new Error("xl/workbook.xml was not found.");
        }
        const sharedStrings = sharedStringsHelper.parseSharedStrings(files);
        const cellStyles = stylesParserHelper.parseCellStyles(files);
        const rels = parseRelationships(files, "xl/_rels/workbook.xml.rels", "xl/workbook.xml");
        const workbookDoc = xmlToDocument(decodeXmlText(workbookBytes));
        const sheetNodes = Array.from(workbookDoc.getElementsByTagName("sheet"));
        const sheetNames = sheetNodes.map((sheetNode, index) => sheetNode.getAttribute("name") || `Sheet${index + 1}`);
        const definedNames = parseDefinedNames(workbookDoc, sheetNames);
        const sheets = sheetNodes.map((sheetNode, index) => {
            const name = sheetNode.getAttribute("name") || `Sheet${index + 1}`;
            const relId = sheetNode.getAttribute("r:id") || "";
            const sheetPath = rels.get(relId) || "";
            return parseWorksheet(files, name, sheetPath, index + 1, sharedStrings, cellStyles);
        });
        const workbook = {
            name: workbookName,
            sheets,
            sharedStrings,
            definedNames
        };
        resolveSimpleFormulaReferences(workbook);
        resolveSimpleFormulaReferences(workbook);
        resolveSimpleFormulaReferences(workbook);
        return workbook;
    }
    function buildCellMap(sheet) {
        const map = new Map();
        for (const cell of sheet.cells) {
            map.set(`${cell.row}:${cell.col}`, cell);
        }
        return map;
    }
    function formatCellForMarkdown(cell, options) {
        if (!cell)
            return "";
        const mode = options.outputMode || "display";
        const displayValue = String(cell.outputValue || "");
        const rawValue = String(cell.rawValue || "");
        if (mode === "raw") {
            return rawValue || displayValue;
        }
        if (mode === "both") {
            if (rawValue && rawValue !== displayValue) {
                if (displayValue) {
                    return `${displayValue} [raw=${rawValue}]`;
                }
                return `[raw=${rawValue}]`;
            }
            return displayValue || rawValue;
        }
        return displayValue;
    }
    function isCellInAnyTable(row, col, tables) {
        return tables.some((table) => row >= table.startRow && row <= table.endRow && col >= table.startCol && col <= table.endCol);
    }
    function extractNarrativeBlocks(sheet, tables, options = {}) {
        const rowMap = new Map();
        for (const cell of sheet.cells) {
            if (!cell.outputValue)
                continue;
            if (isCellInAnyTable(cell.row, cell.col, tables))
                continue;
            const entries = rowMap.get(cell.row) || [];
            entries.push(cell);
            rowMap.set(cell.row, entries);
        }
        const rowNumbers = Array.from(rowMap.keys()).sort((a, b) => a - b);
        const blocks = [];
        let current = null;
        let previousRow = -100;
        for (const rowNumber of rowNumbers) {
            const cells = (rowMap.get(rowNumber) || []).slice().sort((a, b) => a.col - b.col);
            const rowSegments = splitNarrativeRowSegments(cells, options);
            for (const segment of rowSegments) {
                const rowText = segment.values.join(" ").trim();
                if (!rowText)
                    continue;
                const startCol = segment.startCol;
                if (!current || rowNumber - previousRow > 1 || Math.abs(startCol - current.startCol) > 3) {
                    current = {
                        startRow: rowNumber,
                        startCol,
                        endRow: rowNumber,
                        lines: [rowText],
                        items: [{
                                row: rowNumber,
                                startCol,
                                text: rowText,
                                cellValues: segment.values
                            }]
                    };
                    blocks.push(current);
                }
                else {
                    current.lines.push(rowText);
                    current.endRow = rowNumber;
                    current.items.push({
                        row: rowNumber,
                        startCol,
                        text: rowText,
                        cellValues: segment.values
                    });
                }
                previousRow = rowNumber;
            }
        }
        return blocks;
    }
    function splitNarrativeRowSegments(cells, options) {
        const segments = [];
        let current = null;
        for (const cell of cells) {
            const value = formatCellForMarkdown(cell, options).trim();
            if (!value)
                continue;
            if (!current || cell.col - current.lastCol > 4) {
                current = {
                    startCol: cell.col,
                    values: [value],
                    lastCol: cell.col
                };
                segments.push(current);
            }
            else {
                current.values.push(value);
                current.lastCol = cell.col;
            }
        }
        return segments.map((segment) => ({
            startCol: segment.startCol,
            values: segment.values
        }));
    }
    function extractSectionBlocks(sheet, tables, narrativeBlocks) {
        const charts = sheet.charts || [];
        const anchors = [];
        for (const block of narrativeBlocks) {
            anchors.push({
                startRow: block.startRow,
                startCol: block.startCol,
                endRow: block.endRow,
                endCol: Math.max(block.startCol, ...block.items.map((item) => item.startCol))
            });
        }
        for (const table of tables) {
            anchors.push({
                startRow: table.startRow,
                startCol: table.startCol,
                endRow: table.endRow,
                endCol: table.endCol
            });
        }
        for (const image of sheet.images) {
            const anchor = parseCellAddress(image.anchor);
            if (anchor.row > 0 && anchor.col > 0) {
                anchors.push({
                    startRow: anchor.row,
                    startCol: anchor.col,
                    endRow: anchor.row,
                    endCol: anchor.col
                });
            }
        }
        for (const chart of charts) {
            const anchor = parseCellAddress(chart.anchor);
            if (anchor.row > 0 && anchor.col > 0) {
                anchors.push({
                    startRow: anchor.row,
                    startCol: anchor.col,
                    endRow: anchor.row,
                    endCol: anchor.col
                });
            }
        }
        if (anchors.length === 0) {
            return [];
        }
        anchors.sort((left, right) => {
            if (left.startRow !== right.startRow)
                return left.startRow - right.startRow;
            return left.startCol - right.startCol;
        });
        const sections = [];
        let current = null;
        let previousEndRow = -100;
        const verticalGapThreshold = 4;
        for (const anchor of anchors) {
            const gap = anchor.startRow - previousEndRow;
            if (!current || gap > verticalGapThreshold) {
                current = {
                    startRow: anchor.startRow,
                    startCol: anchor.startCol,
                    endRow: anchor.endRow,
                    endCol: anchor.endCol
                };
                sections.push(current);
            }
            else {
                current.startRow = Math.min(current.startRow, anchor.startRow);
                current.startCol = Math.min(current.startCol, anchor.startCol);
                current.endRow = Math.max(current.endRow, anchor.endRow);
                current.endCol = Math.max(current.endCol, anchor.endCol);
            }
            previousEndRow = Math.max(previousEndRow, anchor.endRow);
        }
        return sections;
    }
    function convertSheetToMarkdown(workbook, sheet, options = {}) {
        var _a;
        const charts = sheet.charts || [];
        const shapes = sheet.shapes || [];
        const shapeBlocks = extractShapeBlocks(shapes);
        const treatFirstRowAsHeader = options.treatFirstRowAsHeader !== false;
        const tables = tableDetectorHelper.detectTableCandidates(sheet, buildCellMap);
        const narrativeBlocks = extractNarrativeBlocks(sheet, tables, options);
        const sectionBlocks = extractSectionBlocks(sheet, tables, narrativeBlocks);
        const formulaDiagnostics = sheet.cells
            .filter((cell) => !!cell.formulaText && cell.resolutionStatus !== null)
            .map((cell) => ({
            address: cell.address,
            formulaText: cell.formulaText,
            status: cell.resolutionStatus,
            source: cell.resolutionSource,
            outputValue: cell.outputValue
        }));
        const sections = [];
        for (const block of narrativeBlocks) {
            sections.push({
                sortRow: block.startRow,
                sortCol: block.startCol,
                markdown: `${narrativeStructureHelper.renderNarrativeBlock(block)}\n`,
                kind: "narrative",
                narrativeBlock: block
            });
        }
        let tableCounter = 1;
        for (const table of tables) {
            const rows = tableDetectorHelper.matrixFromCandidate(sheet, table, options, buildCellMap, formatCellForMarkdown);
            if (rows.length === 0 || ((_a = rows[0]) === null || _a === void 0 ? void 0 : _a.length) === 0)
                continue;
            const tableMarkdown = markdownExportHelper.renderMarkdownTable(rows, treatFirstRowAsHeader);
            sections.push({
                sortRow: table.startRow,
                sortCol: table.startCol,
                markdown: `### Table ${String(tableCounter).padStart(3, "0")} (${formatRange(table.startRow, table.startCol, table.endRow, table.endCol)})\n\n${tableMarkdown}\n`,
                kind: "table"
            });
            tableCounter += 1;
        }
        sections.sort((left, right) => {
            if (left.sortRow !== right.sortRow)
                return left.sortRow - right.sortRow;
            return left.sortCol - right.sortCol;
        });
        const groupedSections = (sectionBlocks.length > 0 ? sectionBlocks : [{
                startRow: -1,
                startCol: -1,
                endRow: Number.MAX_SAFE_INTEGER,
                endCol: Number.MAX_SAFE_INTEGER
            }]).map((block) => ({
            block,
            entries: sections.filter((section) => section.sortRow >= block.startRow
                && section.sortRow <= block.endRow
                && section.sortCol >= block.startCol
                && section.sortCol <= block.endCol)
        })).filter((group) => group.entries.length > 0);
        const body = groupedSections
            .map((group) => {
            return group.entries.map((section) => section.markdown.trimEnd()).join("\n\n").trim();
        })
            .filter(Boolean)
            .join("\n\n---\n\n")
            .trim();
        const imageSection = sheet.images.length > 0
            ? [
                "",
                "## Images",
                "",
                ...sheet.images.map((image, index) => [
                    `### Image ${String(index + 1).padStart(3, "0")} (${image.anchor})`,
                    `- File: ${image.path}`,
                    "",
                    `![${image.filename}](${image.path})`
                ].join("\n"))
            ].join("\n")
            : "";
        const chartSection = charts.length > 0
            ? [
                "",
                "## Charts",
                "",
                ...charts.map((chart, index) => {
                    const lines = [
                        `### Chart ${String(index + 1).padStart(3, "0")} (${chart.anchor})`,
                        `- Title: ${chart.title || "(none)"}`,
                        `- Type: ${chart.chartType}`
                    ];
                    if (chart.series.length > 0) {
                        lines.push("- Series:");
                        for (const series of chart.series) {
                            lines.push(`  - ${series.name}`);
                            if (series.axis === "secondary") {
                                lines.push("    - Axis: secondary");
                            }
                            if (series.categoriesRef) {
                                lines.push(`    - categories: ${series.categoriesRef}`);
                            }
                            if (series.valuesRef) {
                                lines.push(`    - values: ${series.valuesRef}`);
                            }
                        }
                    }
                    return lines.join("\n");
                })
            ].join("\n")
            : "";
        const shapeSection = shapes.length > 0
            ? [
                "",
                "## Shape Blocks",
                "",
                ...shapeBlocks.map((block, blockIndex) => [
                    `### Shape Block ${String(blockIndex + 1).padStart(3, "0")} (${formatRange(block.startRow, block.startCol, block.endRow, block.endCol)})`,
                    `- Shapes: ${block.shapeIndexes.map((shapeIndex) => `Shape ${String(shapeIndex + 1).padStart(3, "0")}`).join(", ")}`,
                    `- anchorRange: ${colToLetters(block.startCol)}${block.startRow}-${colToLetters(block.endCol)}${block.endRow}`
                ].join("\n")),
                "",
                "## Shapes",
                "",
                ...shapes.map((shape, index) => {
                    const lines = [
                        `### Shape ${String(index + 1).padStart(3, "0")} (${shape.anchor})`,
                        ...renderHierarchicalRawEntries(shape.rawEntries)
                    ];
                    if (shape.svgPath) {
                        lines.push(`- SVG: ${shape.svgPath}`);
                        lines.push("");
                        lines.push(`![${shape.svgFilename || `shape_${String(index + 1).padStart(3, "0")}.svg`}](${shape.svgPath})`);
                    }
                    return lines.join("\n");
                })
            ].join("\n")
            : "";
        const markdown = [
            `# ${sheet.name}`,
            "",
            "## Source Information",
            `- Workbook: ${workbook.name}`,
            `- Sheet: ${sheet.name}`,
            "",
            "## Body",
            "",
            body || "_No extractable body content was found._",
            chartSection,
            shapeSection,
            imageSection
        ].join("\n");
        return {
            fileName: markdownExportHelper.createOutputFileName(workbook.name, sheet.index, sheet.name, options.outputMode || "display"),
            sheetName: sheet.name,
            markdown,
            summary: {
                outputMode: options.outputMode || "display",
                sections: sectionBlocks.length,
                tables: tables.length,
                narrativeBlocks: narrativeBlocks.length,
                merges: sheet.merges.length,
                images: sheet.images.length,
                charts: charts.length,
                cells: sheet.cells.length,
                tableScores: tables.map((table) => ({
                    range: formatRange(table.startRow, table.startCol, table.endRow, table.endCol),
                    score: table.score,
                    reasons: [...table.reasonSummary]
                })),
                formulaDiagnostics
            }
        };
    }
    function convertWorkbookToMarkdownFiles(workbook, options = {}) {
        return workbook.sheets.map((sheet) => convertSheetToMarkdown(workbook, sheet, options));
    }
    globalThis.__xlsx2md = {
        parseWorkbook,
        unzipEntries: zipIoHelper.unzipEntries,
        parseRangeRef,
        applyMergeTokens: tableDetectorHelper.applyMergeTokens,
        detectTableCandidates: (sheet) => tableDetectorHelper.detectTableCandidates(sheet, buildCellMap),
        extractNarrativeBlocks,
        convertSheetToMarkdown,
        convertWorkbookToMarkdownFiles,
        createSummaryText: markdownExportHelper.createSummaryText,
        createCombinedMarkdownExportFile: markdownExportHelper.createCombinedMarkdownExportFile,
        createExportEntries: markdownExportHelper.createExportEntries,
        createWorkbookExportArchive: markdownExportHelper.createWorkbookExportArchive,
        formatRange,
        colToLetters,
        lettersToCol,
        textEncoder: markdownExportHelper.textEncoder
    };
})();
