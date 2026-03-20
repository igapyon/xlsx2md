(() => {
  const moduleRegistry = getXlsx2mdModuleRegistry();
  type BorderFlags = {
    top: boolean;
    bottom: boolean;
    left: boolean;
    right: boolean;
  };

  type CellStyleInfo = {
    borders: BorderFlags;
    numFmtId: number;
    formatCode: string;
  };

  type FormulaResolutionSource = "cached_value" | "ast_evaluator" | "legacy_resolver" | "formula_text" | "external_unsupported" | null;

  type ResolvedCellLike = {
    borders: BorderFlags;
    numFmtId: number;
    formatCode: string;
    rawValue: string;
    outputValue: string;
    resolutionStatus: "resolved" | "fallback_formula" | "unsupported_external" | null;
    resolutionSource: FormulaResolutionSource;
  };

  function isDateFormatCode(formatCode: string): boolean {
    const normalized = String(formatCode || "")
      .toLowerCase()
      .replace(/\[[^\]]*]/g, "")
      .replace(/"[^"]*"/g, "")
      .replace(/\\./g, "");
    if (!normalized) return false;
    if (normalized.includes("general")) return false;
    return /[ymdhs]/.test(normalized);
  }

  function normalizeNumericFormatCode(formatCode: string): string {
    return String(formatCode || "")
      .trim()
      .replace(/\[[^\]]*]/g, "")
      .replace(/"([^"]*)"/g, "$1")
      .replace(/\\(.)/g, "$1")
      .replace(/_.?/g, "")
      .replace(/\*/g, "");
  }

  function excelSerialToIsoText(serial: number): string {
    if (!Number.isFinite(serial)) return String(serial);
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

  function excelSerialToDateParts(serial: number): {
    year: number;
    month: number;
    day: number;
    hour: number;
    minute: number;
    second: number;
    yyyy: string;
    mm: string;
    dd: string;
    hh: string;
    mi: string;
    ss: string;
  } | null {
    if (!Number.isFinite(serial)) return null;
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

  function formatTextFunctionValue(value: string, formatText: string): string | null {
    const format = String(formatText || "").trim();
    if (!format) return null;
    const numericValue = Number(value);
    const normalized = format.toLowerCase();

    if (!Number.isNaN(numericValue)) {
      if (/(^|[^a-z])yyyy/.test(normalized) || normalized.includes("hh:") || normalized.includes("mm/") || normalized.includes("mm-")) {
        const parts = excelSerialToDateParts(numericValue);
        if (!parts) return null;
        if (normalized === "yyyy-mm-dd") return `${parts.yyyy}-${parts.mm}-${parts.dd}`;
        if (normalized === "yyyy/mm/dd") return `${parts.yyyy}/${parts.mm}/${parts.dd}`;
        if (normalized === "hh:mm:ss") return `${parts.hh}:${parts.mi}:${parts.ss}`;
        if (normalized === "yyyy-mm-dd hh:mm:ss") return `${parts.yyyy}-${parts.mm}-${parts.dd} ${parts.hh}:${parts.mi}:${parts.ss}`;
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

  function formatNumberByPattern(value: number, pattern: string): string {
    const normalizedPattern = pattern.trim();
    const decimalPlaces = (normalizedPattern.split(".")[1] || "").replace(/[^0#]/g, "").length;
    const useGrouping = normalizedPattern.includes(",");
    return value.toLocaleString("en-US", {
      minimumFractionDigits: decimalPlaces,
      maximumFractionDigits: decimalPlaces,
      useGrouping
    });
  }

  function formatDateByPattern(parts: NonNullable<ReturnType<typeof excelSerialToDateParts>>, formatCode: string): string | null {
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

  function formatFractionPattern(value: number): string | null {
    if (!Number.isFinite(value)) return null;
    const tolerance = 1e-9;
    for (let denominator = 1; denominator <= 100; denominator += 1) {
      const numerator = Math.round(value * denominator);
      if (Math.abs(value - (numerator / denominator)) < tolerance) {
        return `${numerator}/${denominator}`;
      }
    }
    return null;
  }

  function formatDbNum3Pattern(rawValue: string): string {
    return rawValue.split("").join(" ");
  }

  function splitFormatSections(formatCode: string): string[] {
    const sections: string[] = [];
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

  function formatZeroSection(section: string): string | null {
    const normalizedSection = String(section || "");
    if (!normalizedSection) return null;
    const compact = normalizedSection.replace(/_.|\\.|[*?]/g, "").trim();
    const hasDashLiteral = /"-"|(^|[^a-z0-9])-($|[^a-z0-9])/i.test(compact);
    if (!hasDashLiteral) return null;
    if (compact.includes("¥")) return "¥ -";
    if (compact.includes("$")) return "$ -";
    return "-";
  }

  function formatCellDisplayValue(rawValue: string, cellStyle: CellStyleInfo): string | null {
    if (rawValue === "") return null;
    const numericValue = Number(rawValue);
    const formatCode = normalizeNumericFormatCode(cellStyle.formatCode);
    const normalized = formatCode.toLowerCase();
    const formatSections = splitFormatSections(formatCode);

    if (!Number.isNaN(numericValue) && isDateFormatCode(formatCode)) {
      const parts = excelSerialToDateParts(numericValue);
      if (!parts) return null;
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
      const decimalPlaces = (decimalPartMatch?.[1] || "").length;
      const exponentDigits = (scientificPattern.match(/e\+([0#]+)/i)?.[1] || "").length;
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

  function applyResolvedFormulaValue(
    cell: ResolvedCellLike,
    resolvedValue: string,
    resolutionSource: FormulaResolutionSource = "legacy_resolver"
  ): void {
    const rawValue = String(resolvedValue || "");
    const formattedValue = formatCellDisplayValue(rawValue, {
      borders: cell.borders,
      numFmtId: cell.numFmtId,
      formatCode: cell.formatCode
    });
    cell.rawValue = rawValue;
    cell.outputValue = formattedValue ?? rawValue;
    cell.resolutionStatus = "resolved";
    cell.resolutionSource = resolutionSource;
  }

  function parseDateLikeParts(value: string): {
    yyyy: string;
    mm: string;
    dd: string;
    hh: string;
    mi: string;
    ss: string;
  } | null {
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

  function datePartsToExcelSerial(
    year: number,
    month: number,
    day: number,
    hour = 0,
    minute = 0,
    second = 0
  ): number | null {
    if (![year, month, day, hour, minute, second].every(Number.isFinite)) return null;
    const baseUtcMs = Date.UTC(1899, 11, 31);
    const targetUtcMs = Date.UTC(year, month - 1, day, hour, minute, second);
    const msPerDay = 24 * 60 * 60 * 1000;
    let serial = (targetUtcMs - baseUtcMs) / msPerDay;
    if (serial >= 60) {
      serial += 1;
    }
    return serial;
  }

  function parseValueFunctionText(value: string): number | null {
    const trimmed = String(value || "").trim();
    if (!trimmed) return null;
    const numericValue = Number(trimmed.replace(/,/g, ""));
    if (!Number.isNaN(numericValue)) {
      return numericValue;
    }
    const parts = parseDateLikeParts(trimmed);
    if (!parts) return null;
    return datePartsToExcelSerial(
      Number(parts.yyyy),
      Number(parts.mm),
      Number(parts.dd),
      Number(parts.hh),
      Number(parts.mi),
      Number(parts.ss)
    );
  }

  const cellFormatApi = {
    isDateFormatCode,
    normalizeNumericFormatCode,
    excelSerialToIsoText,
    excelSerialToDateParts,
    formatTextFunctionValue,
    formatNumberByPattern,
    formatDateByPattern,
    formatFractionPattern,
    formatDbNum3Pattern,
    splitFormatSections,
    formatZeroSection,
    formatCellDisplayValue,
    applyResolvedFormulaValue,
    parseDateLikeParts,
    datePartsToExcelSerial,
    parseValueFunctionText
  };

  moduleRegistry.registerModule("cellFormat", cellFormatApi);
})();
