(function initXlsx2mdFormulaEvaluator(global) {
    const moduleRegistry = getXlsx2mdModuleRegistry();
    const api = moduleRegistry.getModule("formulaRuntime");
    if (!api) {
        throw new Error("xlsx2md formula runtime module is not loaded");
    }
    function evaluateFormulaAst(ast, context = {}) {
        switch (ast.type) {
            case "number":
                return ast.value;
            case "string":
                return ast.value;
            case "boolean":
                return ast.value;
            case "error":
                return ast.value;
            case "array_constant":
                return ast.rows.map((row) => row.map((item) => evaluateFormulaAst(item, context)));
            case "cell":
                return context.resolveCell ? context.resolveCell(ast.ref, ast.sheet) : null;
            case "name":
                return context.resolveName ? context.resolveName(ast.name) : null;
            case "scoped_name":
                if (context.resolveScopedName) {
                    return context.resolveScopedName(ast.sheet, ast.name);
                }
                return context.resolveName ? context.resolveName(`${ast.sheet}!${ast.name}`) : null;
            case "range":
                return evaluateRangeAst(ast, context);
            case "structured_ref":
                if (ast.qualifier) {
                    return null;
                }
                return context.resolveStructuredRef ? context.resolveStructuredRef(ast.table, ast.column) : null;
            case "unary_op":
                return evaluateUnaryOp(ast.operator, evaluateFormulaAst(ast.operand, context));
            case "postfix_op":
                if (ast.operator === "#" && ast.operand.type === "cell") {
                    return context.resolveSpill ? context.resolveSpill(ast.operand.ref, ast.operand.sheet) : null;
                }
                return evaluatePostfixOp(ast.operator, evaluateFormulaAst(ast.operand, context));
            case "binary_op":
                if (ast.operator === " ") {
                    return evaluateIntersectionAst(ast, context);
                }
                return evaluateBinaryOp(ast.operator, evaluateFormulaAst(ast.left, context), evaluateFormulaAst(ast.right, context));
            case "function_call":
                return evaluateFunctionCall(ast.name, ast.args, context);
            default:
                throw new Error(`Unsupported AST node: ${ast.type}`);
        }
    }
    function evaluateRangeAst(ast, context) {
        var _a, _b;
        if (ast.start.type === "cell" && ast.end.type === "cell") {
            const sheet = (_b = (_a = ast.start.sheet) !== null && _a !== void 0 ? _a : ast.end.sheet) !== null && _b !== void 0 ? _b : null;
            if (context.resolveRange) {
                return context.resolveRange(ast.start.ref, ast.end.ref, sheet);
            }
            return [
                evaluateFormulaAst(ast.start, context),
                evaluateFormulaAst(ast.end, context)
            ];
        }
        return [
            evaluateFormulaAst(ast.start, context),
            evaluateFormulaAst(ast.end, context)
        ];
    }
    function evaluateIntersectionAst(ast, context) {
        var _a, _b, _c, _d;
        const leftArea = toCellArea(ast.left);
        const rightArea = toCellArea(ast.right);
        if (!leftArea || !rightArea) {
            throw new Error("Unsupported intersection operands");
        }
        const leftSheet = (_b = (_a = leftArea.sheet) !== null && _a !== void 0 ? _a : rightArea.sheet) !== null && _b !== void 0 ? _b : null;
        const rightSheet = (_d = (_c = rightArea.sheet) !== null && _c !== void 0 ? _c : leftArea.sheet) !== null && _d !== void 0 ? _d : null;
        if (leftSheet !== rightSheet) {
            return "#NULL!";
        }
        const startRow = Math.max(leftArea.startRow, rightArea.startRow);
        const endRow = Math.min(leftArea.endRow, rightArea.endRow);
        const startCol = Math.max(leftArea.startCol, rightArea.startCol);
        const endCol = Math.min(leftArea.endCol, rightArea.endCol);
        if (startRow > endRow || startCol > endCol) {
            return "#NULL!";
        }
        const startRef = `${colToLetters(startCol)}${startRow}`;
        const endRef = `${colToLetters(endCol)}${endRow}`;
        if (context.resolveRange) {
            return context.resolveRange(startRef, endRef, leftSheet);
        }
        return [[`${startRef}:${endRef}`]];
    }
    function evaluateUnaryOp(operator, operand) {
        const numericValue = toNumber(operand);
        if (operator === "+") {
            return numericValue;
        }
        if (operator === "-") {
            return -numericValue;
        }
        throw new Error(`Unsupported unary operator: ${operator}`);
    }
    function evaluatePostfixOp(operator, operand) {
        if (operator === "%") {
            return toNumber(operand) / 100;
        }
        throw new Error(`Unsupported postfix operator: ${operator}`);
    }
    function evaluateBinaryOp(operator, left, right) {
        switch (operator) {
            case "+":
                return toNumber(left) + toNumber(right);
            case "-":
                return toNumber(left) - toNumber(right);
            case "*":
                return toNumber(left) * toNumber(right);
            case "/":
                return toNumber(left) / toNumber(right);
            case "&":
                return `${toText(left)}${toText(right)}`;
            case "=":
                return looselyEquals(left, right);
            case "<>":
                return !looselyEquals(left, right);
            case "<":
                return compareValues(left, right) < 0;
            case "<=":
                return compareValues(left, right) <= 0;
            case ">":
                return compareValues(left, right) > 0;
            case ">=":
                return compareValues(left, right) >= 0;
            default:
                throw new Error(`Unsupported binary operator: ${operator}`);
        }
    }
    function toCellArea(node) {
        var _a, _b;
        if (node.type === "cell") {
            const position = parseCellRef(node.ref);
            if (!position) {
                return null;
            }
            return {
                sheet: node.sheet,
                startRow: position.row,
                endRow: position.row,
                startCol: position.col,
                endCol: position.col
            };
        }
        if (node.type === "range" && node.start.type === "cell" && node.end.type === "cell") {
            const start = parseCellRef(node.start.ref);
            const end = parseCellRef(node.end.ref);
            if (!start || !end) {
                return null;
            }
            return {
                sheet: (_b = (_a = node.start.sheet) !== null && _a !== void 0 ? _a : node.end.sheet) !== null && _b !== void 0 ? _b : null,
                startRow: Math.min(start.row, end.row),
                endRow: Math.max(start.row, end.row),
                startCol: Math.min(start.col, end.col),
                endCol: Math.max(start.col, end.col)
            };
        }
        return null;
    }
    function parseCellRef(ref) {
        const match = String(ref).toUpperCase().match(/^\$?([A-Z]{1,3})\$?(\d+)$/);
        if (!match) {
            return null;
        }
        return {
            col: lettersToCol(match[1]),
            row: Number(match[2])
        };
    }
    function lettersToCol(letters) {
        let value = 0;
        for (const char of letters) {
            value = value * 26 + (char.charCodeAt(0) - 64);
        }
        return value;
    }
    function colToLetters(column) {
        let current = column;
        let result = "";
        while (current > 0) {
            const remainder = (current - 1) % 26;
            result = String.fromCharCode(65 + remainder) + result;
            current = Math.floor((current - 1) / 26);
        }
        return result;
    }
    function evaluateFunctionCall(name, args, context) {
        const upperName = name.toUpperCase();
        switch (upperName) {
            case "IF":
                return evaluateIf(args, context);
            case "IFERROR":
                return evaluateIfError(args, context);
            case "AND":
                return evaluateAnd(args, context);
            case "OR":
                return evaluateOr(args, context);
            case "NOT":
                return evaluateNot(args, context);
            case "DATE":
                return evaluateDate(args, context);
            case "VALUE":
                return evaluateValue(args, context);
            case "ROUND":
                return evaluateRound(args, context, "round");
            case "ROUNDUP":
                return evaluateRound(args, context, "up");
            case "ROUNDDOWN":
                return evaluateRound(args, context, "down");
            case "INT":
                return evaluateInt(args, context);
            case "ABS":
                return evaluateAbs(args, context);
            case "SUM":
                return evaluateSum(args, context);
            case "SUMPRODUCT":
                return evaluateSumProduct(args, context);
            case "REPT":
                return evaluateRept(args, context);
            case "SUBSTITUTE":
                return evaluateSubstitute(args, context);
            case "MATCH":
                return evaluateMatch(args, context);
            case "INDEX":
                return evaluateIndex(args, context);
            case "VLOOKUP":
                return evaluateVLookup(args, context);
            case "HLOOKUP":
                return evaluateHLookup(args, context);
            case "XLOOKUP":
                return evaluateXLookup(args, context);
            case "TEXT":
                return evaluateText(args, context);
            case "TODAY":
                return evaluateToday(context);
            case "WEEKDAY":
                return evaluateWeekday(args, context);
            case "DATEVALUE":
                return evaluateDateValue(args, context);
            case "LEN":
                return evaluateLen(args, context);
            case "LOWER":
                return evaluateLower(args, context);
            case "FIND":
                return evaluateFind(args, context, false);
            case "SEARCH":
                return evaluateFind(args, context, true);
            case "LEFT":
                return evaluateLeft(args, context);
            case "RIGHT":
                return evaluateRight(args, context);
            case "MID":
                return evaluateMid(args, context);
            case "TRIM":
                return evaluateTrim(args, context);
            case "REPLACE":
                return evaluateReplace(args, context);
            case "DAY":
                return evaluateDay(args, context);
            case "MONTH":
                return evaluateMonth(args, context);
            case "YEAR":
                return evaluateYear(args, context);
            case "SUBTOTAL":
                return evaluateSubtotal(args, context);
            case "UPPER":
                return evaluateUpper(args, context);
            case "CONCATENATE":
                return evaluateConcatenate(args, context);
            case "ISBLANK":
                return evaluateIsBlank(args, context);
            case "ISNUMBER":
                return evaluateIsNumber(args, context);
            case "ISTEXT":
                return evaluateIsText(args, context);
            case "ISERROR":
                return evaluateIsError(args, context);
            case "ISNA":
                return evaluateIsNa(args, context);
            case "NA":
                return evaluateNa();
            case "MIN":
                return evaluateMin(args, context);
            case "MAX":
                return evaluateMax(args, context);
            case "AVERAGE":
                return evaluateAverage(args, context);
            case "COLUMN":
                return evaluateColumn(args, context);
            case "ROW":
                return evaluateRow(args, context);
            case "EDATE":
                return evaluateEDate(args, context);
            case "EOMONTH":
                return evaluateEoMonth(args, context);
            case "COUNTIF":
                return evaluateCountIf(args, context);
            case "COUNTIFS":
                return evaluateCountIfs(args, context);
            case "COUNT":
                return evaluateCount(args, context);
            case "COUNTA":
                return evaluateCountA(args, context);
            case "SUMIF":
                return evaluateSumIf(args, context);
            case "SUMIFS":
                return evaluateSumIfs(args, context);
            case "AVERAGEIF":
                return evaluateAverageIf(args, context);
            case "AVERAGEIFS":
                return evaluateAverageIfs(args, context);
            default:
                throw new Error(`Unsupported formula function: ${name}`);
        }
    }
    function evaluateIf(args, context) {
        const condition = toBoolean(evaluateFormulaAst(args[0], context));
        if (condition) {
            return args[1] ? evaluateFormulaAst(args[1], context) : true;
        }
        return args[2] ? evaluateFormulaAst(args[2], context) : false;
    }
    function evaluateIfError(args, context) {
        const primary = evaluateFormulaAst(args[0], context);
        if (isFormulaError(primary)) {
            return args[1] ? evaluateFormulaAst(args[1], context) : "";
        }
        return primary;
    }
    function evaluateAnd(args, context) {
        return args.every((arg) => toBoolean(evaluateFormulaAst(arg, context)));
    }
    function evaluateOr(args, context) {
        return args.some((arg) => toBoolean(evaluateFormulaAst(arg, context)));
    }
    function evaluateNot(args, context) {
        return !toBoolean(evaluateFormulaAst(args[0], context));
    }
    function evaluateDate(args, context) {
        const year = toNumber(evaluateFormulaAst(args[0], context));
        const month = toNumber(evaluateFormulaAst(args[1], context));
        const day = toNumber(evaluateFormulaAst(args[2], context));
        return excelSerialFromDate(year, month, day);
    }
    function evaluateValue(args, context) {
        const rawValue = evaluateFormulaAst(args[0], context);
        if (typeof rawValue === "number") {
            return rawValue;
        }
        const text = toText(rawValue).trim();
        if (!text) {
            return 0;
        }
        const dateValue = parseDateLikeString(text);
        if (dateValue !== null) {
            return dateValue;
        }
        const normalized = text.replace(/,/g, "");
        const parsed = Number(normalized);
        if (!Number.isNaN(parsed)) {
            return parsed;
        }
        throw new Error(`Unsupported VALUE input: ${text}`);
    }
    function evaluateRound(args, context, mode) {
        const value = toNumber(evaluateFormulaAst(args[0], context));
        const digits = args[1] ? Math.trunc(toNumber(evaluateFormulaAst(args[1], context))) : 0;
        const factor = Math.pow(10, digits);
        const scaled = value * factor;
        if (mode === "round") {
            return Math.round(scaled) / factor;
        }
        if (mode === "up") {
            return (scaled >= 0 ? Math.ceil(scaled) : Math.floor(scaled)) / factor;
        }
        return (scaled >= 0 ? Math.floor(scaled) : Math.ceil(scaled)) / factor;
    }
    function evaluateInt(args, context) {
        return Math.floor(toNumber(evaluateFormulaAst(args[0], context)));
    }
    function evaluateAbs(args, context) {
        return Math.abs(toNumber(evaluateFormulaAst(args[0], context)));
    }
    function evaluateSum(args, context) {
        return args
            .flatMap((arg) => flattenValues(evaluateFormulaAst(arg, context)))
            .reduce((sum, value) => sum + toNumber(value), 0);
    }
    function evaluateSumProduct(args, context) {
        var _a;
        const vectors = args.map((arg) => flattenValues(evaluateFormulaAst(arg, context)));
        if (!vectors.length) {
            return 0;
        }
        const lengths = vectors.map((vector) => vector.length);
        const maxLength = Math.max(...lengths);
        const normalized = vectors.map((vector) => {
            if (vector.length === maxLength) {
                return vector;
            }
            if (vector.length === 1) {
                return Array.from({ length: maxLength }, () => vector[0]);
            }
            throw new Error("SUMPRODUCT arguments must have the same length");
        });
        let total = 0;
        for (let index = 0; index < maxLength; index += 1) {
            let product = 1;
            for (const vector of normalized) {
                product *= toNumber((_a = vector[index]) !== null && _a !== void 0 ? _a : 0);
            }
            total += product;
        }
        return total;
    }
    function evaluateRept(args, context) {
        const text = toText(evaluateFormulaAst(args[0], context));
        const countValue = evaluateFormulaAst(args[1], context);
        const count = Math.max(0, Math.floor(toNumber(countValue)));
        return text.repeat(count);
    }
    function evaluateSubstitute(args, context) {
        const text = toText(evaluateFormulaAst(args[0], context));
        const oldText = toText(evaluateFormulaAst(args[1], context));
        const newText = toText(evaluateFormulaAst(args[2], context));
        const instanceNum = args[3] ? Math.floor(toNumber(evaluateFormulaAst(args[3], context))) : null;
        if (!oldText) {
            return text;
        }
        if (!instanceNum || instanceNum < 1) {
            return text.split(oldText).join(newText);
        }
        let occurrence = 0;
        let searchIndex = 0;
        let result = "";
        while (true) {
            const foundIndex = text.indexOf(oldText, searchIndex);
            if (foundIndex === -1) {
                result += text.slice(searchIndex);
                break;
            }
            occurrence += 1;
            result += text.slice(searchIndex, foundIndex);
            if (occurrence === instanceNum) {
                result += newText;
            }
            else {
                result += oldText;
            }
            searchIndex = foundIndex + oldText.length;
        }
        return result;
    }
    function evaluateMatch(args, context) {
        const lookupValue = evaluateFormulaAst(args[0], context);
        const lookupArray = flattenValues(evaluateFormulaAst(args[1], context));
        for (let index = 0; index < lookupArray.length; index += 1) {
            if (looselyEquals(lookupArray[index], lookupValue)) {
                return index + 1;
            }
        }
        return "#N/A";
    }
    function evaluateIndex(args, context) {
        var _a, _b, _c, _d;
        const source = evaluateFormulaAst(args[0], context);
        const rowNumber = args[1] ? Math.max(1, Math.floor(toNumber(evaluateFormulaAst(args[1], context)))) : 1;
        const columnNumber = args[2] ? Math.max(1, Math.floor(toNumber(evaluateFormulaAst(args[2], context)))) : 1;
        if (!Array.isArray(source)) {
            return source;
        }
        if (source.length > 0 && Array.isArray(source[0])) {
            const row = (_a = source[rowNumber - 1]) !== null && _a !== void 0 ? _a : [];
            return (_b = row[columnNumber - 1]) !== null && _b !== void 0 ? _b : null;
        }
        if (columnNumber === 1) {
            return (_c = source[rowNumber - 1]) !== null && _c !== void 0 ? _c : null;
        }
        return (_d = source[columnNumber - 1]) !== null && _d !== void 0 ? _d : null;
    }
    function evaluateVLookup(args, context) {
        var _a, _b, _c;
        const lookupValue = evaluateFormulaAst(args[0], context);
        const table = normalizeToMatrix(evaluateFormulaAst(args[1], context));
        const columnNumber = Math.max(1, Math.floor(toNumber(evaluateFormulaAst(args[2], context))));
        const approximate = args[3] ? toBoolean(evaluateFormulaAst(args[3], context)) : true;
        if (approximate) {
            let matchedRow = null;
            for (const row of table) {
                if (looselyEquals(row[0], lookupValue)) {
                    return (_a = row[columnNumber - 1]) !== null && _a !== void 0 ? _a : "#N/A";
                }
                if (compareValues(row[0], lookupValue) <= 0) {
                    matchedRow = row;
                }
            }
            return matchedRow ? (_b = matchedRow[columnNumber - 1]) !== null && _b !== void 0 ? _b : "#N/A" : "#N/A";
        }
        for (const row of table) {
            if (looselyEquals(row[0], lookupValue)) {
                return (_c = row[columnNumber - 1]) !== null && _c !== void 0 ? _c : "#N/A";
            }
        }
        return "#N/A";
    }
    function evaluateHLookup(args, context) {
        var _a, _b, _c, _d, _e;
        const lookupValue = evaluateFormulaAst(args[0], context);
        const table = normalizeToMatrix(evaluateFormulaAst(args[1], context));
        const rowNumber = Math.max(1, Math.floor(toNumber(evaluateFormulaAst(args[2], context))));
        const approximate = args[3] ? toBoolean(evaluateFormulaAst(args[3], context)) : true;
        const headerRow = (_a = table[0]) !== null && _a !== void 0 ? _a : [];
        const targetRow = (_b = table[rowNumber - 1]) !== null && _b !== void 0 ? _b : [];
        if (approximate) {
            let matchedIndex = -1;
            for (let index = 0; index < headerRow.length; index += 1) {
                if (looselyEquals(headerRow[index], lookupValue)) {
                    return (_c = targetRow[index]) !== null && _c !== void 0 ? _c : "#N/A";
                }
                if (compareValues(headerRow[index], lookupValue) <= 0) {
                    matchedIndex = index;
                }
            }
            return matchedIndex >= 0 ? (_d = targetRow[matchedIndex]) !== null && _d !== void 0 ? _d : "#N/A" : "#N/A";
        }
        for (let index = 0; index < headerRow.length; index += 1) {
            if (looselyEquals(headerRow[index], lookupValue)) {
                return (_e = targetRow[index]) !== null && _e !== void 0 ? _e : "#N/A";
            }
        }
        return "#N/A";
    }
    function evaluateXLookup(args, context) {
        var _a, _b, _c, _d, _e;
        const lookupValue = evaluateFormulaAst(args[0], context);
        const lookupArray = flattenValues(evaluateFormulaAst(args[1], context));
        const returnArray = flattenValues(evaluateFormulaAst(args[2], context));
        const notFoundValue = args[3] ? evaluateFormulaAst(args[3], context) : "#N/A";
        const matchMode = args[4] ? Math.trunc(toNumber(evaluateFormulaAst(args[4], context))) : 0;
        const searchMode = args[5] ? Math.trunc(toNumber(evaluateFormulaAst(args[5], context))) : 1;
        if (searchMode === 2 || searchMode === -2) {
            const matchedIndex = findXLookupBinaryIndex(lookupArray, lookupValue, matchMode, searchMode);
            return matchedIndex >= 0 ? (_a = returnArray[matchedIndex]) !== null && _a !== void 0 ? _a : notFoundValue : notFoundValue;
        }
        const indices = searchMode === -1
            ? Array.from({ length: lookupArray.length }, (_, index) => lookupArray.length - 1 - index)
            : Array.from({ length: lookupArray.length }, (_, index) => index);
        for (const index of indices) {
            if (looselyEquals(lookupArray[index], lookupValue)) {
                return (_b = returnArray[index]) !== null && _b !== void 0 ? _b : notFoundValue;
            }
        }
        if (matchMode === 2) {
            const matcher = createExcelWildcardMatcher(lookupValue);
            if (!matcher) {
                return notFoundValue;
            }
            for (const index of indices) {
                if (matcher(toText(lookupArray[index]))) {
                    return (_c = returnArray[index]) !== null && _c !== void 0 ? _c : notFoundValue;
                }
            }
            return notFoundValue;
        }
        if (matchMode === -1) {
            let matchedIndex = -1;
            for (const index of indices) {
                if (compareValues(lookupArray[index], lookupValue) <= 0) {
                    matchedIndex = index;
                    if (searchMode === -1) {
                        break;
                    }
                }
            }
            return matchedIndex >= 0 ? (_d = returnArray[matchedIndex]) !== null && _d !== void 0 ? _d : notFoundValue : notFoundValue;
        }
        if (matchMode === 1) {
            let matchedIndex = -1;
            for (const index of indices) {
                if (compareValues(lookupArray[index], lookupValue) >= 0) {
                    matchedIndex = index;
                    break;
                }
            }
            return matchedIndex >= 0 ? (_e = returnArray[matchedIndex]) !== null && _e !== void 0 ? _e : notFoundValue : notFoundValue;
        }
        return notFoundValue;
    }
    function findXLookupBinaryIndex(lookupArray, lookupValue, matchMode, searchMode) {
        const descending = searchMode === -2;
        let low = 0;
        let high = lookupArray.length - 1;
        let fallbackIndex = -1;
        while (low <= high) {
            const mid = Math.floor((low + high) / 2);
            const compare = compareValues(lookupArray[mid], lookupValue);
            if (looselyEquals(lookupArray[mid], lookupValue)) {
                return mid;
            }
            if (matchMode === -1) {
                if (compare <= 0 && (fallbackIndex < 0 || compareValues(lookupArray[mid], lookupArray[fallbackIndex]) > 0)) {
                    fallbackIndex = mid;
                }
            }
            else if (matchMode === 1) {
                if (compare >= 0 && (fallbackIndex < 0 || compareValues(lookupArray[mid], lookupArray[fallbackIndex]) < 0)) {
                    fallbackIndex = mid;
                }
            }
            if ((!descending && compare < 0) || (descending && compare > 0)) {
                low = mid + 1;
            }
            else {
                high = mid - 1;
            }
        }
        return fallbackIndex;
    }
    function evaluateText(args, context) {
        const value = evaluateFormulaAst(args[0], context);
        const format = toText(evaluateFormulaAst(args[1], context)).toLowerCase();
        if (format === "0000") {
            const number = Math.floor(Math.abs(toNumber(value)));
            const sign = toNumber(value) < 0 ? "-" : "";
            return `${sign}${String(number).padStart(4, "0")}`;
        }
        if (format === "0" || format === "0.0" || format === "0.00") {
            const digits = format.includes(".") ? format.split(".")[1].length : 0;
            return toNumber(value).toFixed(digits);
        }
        if (format === "#,##0" || format === "#,##0.00") {
            const digits = format.includes(".") ? format.split(".")[1].length : 0;
            return toNumber(value).toLocaleString("en-US", {
                minimumFractionDigits: digits,
                maximumFractionDigits: digits
            });
        }
        if (format === "yyyy/mm/dd" || format === "yyyy-mm-dd") {
            const parts = excelSerialToDateParts(toNumber(value));
            const separator = format.includes("/") ? "/" : "-";
            return `${parts.year}${separator}${parts.month}${separator}${parts.day}`;
        }
        return toText(value);
    }
    function createExcelWildcardMatcher(patternValue) {
        const pattern = toText(patternValue);
        if (!pattern) {
            return null;
        }
        let regexText = "^";
        for (let index = 0; index < pattern.length; index += 1) {
            const char = pattern[index];
            if (char === "~" && index + 1 < pattern.length) {
                regexText += escapeRegExp(pattern[index + 1]);
                index += 1;
                continue;
            }
            if (char === "*") {
                regexText += ".*";
                continue;
            }
            if (char === "?") {
                regexText += ".";
                continue;
            }
            regexText += escapeRegExp(char);
        }
        regexText += "$";
        const regex = new RegExp(regexText, "i");
        return (value) => regex.test(String(value !== null && value !== void 0 ? value : ""));
    }
    function escapeRegExp(value) {
        return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    }
    function evaluateToday(context) {
        var _a;
        const now = (_a = context.currentDate) !== null && _a !== void 0 ? _a : new Date();
        return excelSerialFromDate(now.getUTCFullYear(), now.getUTCMonth() + 1, now.getUTCDate());
    }
    function evaluateWeekday(args, context) {
        const serial = toNumber(evaluateFormulaAst(args[0], context));
        const returnType = args[1] ? Math.max(1, Math.floor(toNumber(evaluateFormulaAst(args[1], context)))) : 1;
        const excelEpoch = Date.UTC(1899, 11, 30);
        const utcDate = new Date(excelEpoch + Math.floor(serial) * 86400000);
        const jsDay = utcDate.getUTCDay(); // 0=Sun..6=Sat
        switch (returnType) {
            case 1:
                return jsDay + 1;
            case 2:
                return jsDay === 0 ? 7 : jsDay;
            case 3:
                return jsDay === 0 ? 6 : jsDay - 1;
            default:
                return jsDay + 1;
        }
    }
    function evaluateDateValue(args, context) {
        const text = toText(evaluateFormulaAst(args[0], context)).trim();
        const dateValue = parseDateLikeString(text);
        if (dateValue === null) {
            throw new Error(`Unsupported DATEVALUE input: ${text}`);
        }
        return dateValue;
    }
    function evaluateLen(args, context) {
        return toText(evaluateFormulaAst(args[0], context)).length;
    }
    function evaluateLower(args, context) {
        return toText(evaluateFormulaAst(args[0], context)).toLowerCase();
    }
    function evaluateFind(args, context, ignoreCase) {
        const findTextRaw = toText(evaluateFormulaAst(args[0], context));
        const withinTextRaw = toText(evaluateFormulaAst(args[1], context));
        const start = args[2] ? Math.max(1, Math.floor(toNumber(evaluateFormulaAst(args[2], context)))) : 1;
        const findText = ignoreCase ? findTextRaw.toLowerCase() : findTextRaw;
        const withinText = ignoreCase ? withinTextRaw.toLowerCase() : withinTextRaw;
        const index = withinText.indexOf(findText, start - 1);
        return index === -1 ? "#VALUE!" : index + 1;
    }
    function evaluateLeft(args, context) {
        const text = toText(evaluateFormulaAst(args[0], context));
        const count = args[1] ? Math.max(0, Math.floor(toNumber(evaluateFormulaAst(args[1], context)))) : 1;
        return text.slice(0, count);
    }
    function evaluateRight(args, context) {
        const text = toText(evaluateFormulaAst(args[0], context));
        const count = args[1] ? Math.max(0, Math.floor(toNumber(evaluateFormulaAst(args[1], context)))) : 1;
        return count === 0 ? "" : text.slice(-count);
    }
    function evaluateMid(args, context) {
        const text = toText(evaluateFormulaAst(args[0], context));
        const start = Math.max(1, Math.floor(toNumber(evaluateFormulaAst(args[1], context))));
        const count = Math.max(0, Math.floor(toNumber(evaluateFormulaAst(args[2], context))));
        return text.slice(start - 1, start - 1 + count);
    }
    function evaluateTrim(args, context) {
        const text = toText(evaluateFormulaAst(args[0], context));
        return text.trim().replace(/\s+/g, " ");
    }
    function evaluateReplace(args, context) {
        const text = toText(evaluateFormulaAst(args[0], context));
        const start = Math.max(1, Math.floor(toNumber(evaluateFormulaAst(args[1], context))));
        const count = Math.max(0, Math.floor(toNumber(evaluateFormulaAst(args[2], context))));
        const newText = toText(evaluateFormulaAst(args[3], context));
        const prefix = text.slice(0, start - 1);
        const suffix = text.slice(start - 1 + count);
        return `${prefix}${newText}${suffix}`;
    }
    function evaluateDay(args, context) {
        const serial = coerceDateSerial(evaluateFormulaAst(args[0], context));
        return excelSerialToDateParts(serial).day;
    }
    function evaluateMonth(args, context) {
        const serial = coerceDateSerial(evaluateFormulaAst(args[0], context));
        return excelSerialToDateParts(serial).month;
    }
    function evaluateYear(args, context) {
        const serial = coerceDateSerial(evaluateFormulaAst(args[0], context));
        return excelSerialToDateParts(serial).year;
    }
    function evaluateSubtotal(args, context) {
        const functionNum = Math.floor(toNumber(evaluateFormulaAst(args[0], context)));
        const values = args.slice(1).flatMap((arg) => flattenValues(evaluateFormulaAst(arg, context)));
        switch (functionNum) {
            case 1:
            case 101:
                return values.length ? values.reduce((sum, value) => sum + toNumber(value), 0) / values.length : "#DIV/0!";
            case 4:
            case 104:
                return values.reduce((max, value) => Math.max(max, toNumber(value)), Number.NEGATIVE_INFINITY);
            case 5:
            case 105:
                return values.reduce((min, value) => Math.min(min, toNumber(value)), Number.POSITIVE_INFINITY);
            case 9:
            case 109:
                return values.reduce((sum, value) => sum + toNumber(value), 0);
            default:
                throw new Error(`Unsupported SUBTOTAL function_num: ${functionNum}`);
        }
    }
    function evaluateUpper(args, context) {
        return toText(evaluateFormulaAst(args[0], context)).toUpperCase();
    }
    function evaluateConcatenate(args, context) {
        return args.map((arg) => toText(evaluateFormulaAst(arg, context))).join("");
    }
    function evaluateIsBlank(args, context) {
        const value = evaluateFormulaAst(args[0], context);
        return value === null || value === undefined || value === "";
    }
    function evaluateIsNumber(args, context) {
        const value = evaluateFormulaAst(args[0], context);
        if (typeof value === "number") {
            return true;
        }
        if (typeof value === "string") {
            if (!value.trim()) {
                return false;
            }
            const parsed = Number(value.replace(/,/g, ""));
            return !Number.isNaN(parsed);
        }
        return false;
    }
    function evaluateIsText(args, context) {
        const value = evaluateFormulaAst(args[0], context);
        return typeof value === "string";
    }
    function evaluateIsError(args, context) {
        return isFormulaError(evaluateFormulaAst(args[0], context));
    }
    function evaluateIsNa(args, context) {
        return evaluateFormulaAst(args[0], context) === "#N/A";
    }
    function evaluateNa() {
        return "#N/A";
    }
    function evaluateMin(args, context) {
        const values = args.flatMap((arg) => flattenValues(evaluateFormulaAst(arg, context))).map((value) => toNumber(value));
        return values.length ? Math.min(...values) : 0;
    }
    function evaluateMax(args, context) {
        const values = args.flatMap((arg) => flattenValues(evaluateFormulaAst(arg, context))).map((value) => toNumber(value));
        return values.length ? Math.max(...values) : 0;
    }
    function evaluateAverage(args, context) {
        const values = args.flatMap((arg) => flattenValues(evaluateFormulaAst(arg, context))).map((value) => toNumber(value));
        if (!values.length) {
            return "#DIV/0!";
        }
        return values.reduce((sum, value) => sum + value, 0) / values.length;
    }
    function evaluateColumn(args, context) {
        if (!args.length) {
            if (context.currentCellRef) {
                return columnNumberFromRef(context.currentCellRef);
            }
            throw new Error("COLUMN without explicit reference is not supported");
        }
        const node = args[0];
        if (node.type === "cell") {
            return columnNumberFromRef(node.ref);
        }
        if (node.type === "range" && node.start.type === "cell") {
            return columnNumberFromRef(node.start.ref);
        }
        const value = evaluateFormulaAst(node, context);
        if (typeof value === "string" && /\$?[A-Za-z]{1,3}\$?\d+/.test(value)) {
            return columnNumberFromRef(value);
        }
        throw new Error("Unsupported COLUMN argument");
    }
    function evaluateRow(args, context) {
        if (!args.length) {
            if (context.currentCellRef) {
                return rowNumberFromRef(context.currentCellRef);
            }
            throw new Error("ROW without explicit reference is not supported");
        }
        const node = args[0];
        if (node.type === "cell") {
            return rowNumberFromRef(node.ref);
        }
        if (node.type === "range" && node.start.type === "cell") {
            return rowNumberFromRef(node.start.ref);
        }
        const value = evaluateFormulaAst(node, context);
        if (typeof value === "string" && /\$?[A-Za-z]{1,3}\$?\d+/.test(value)) {
            return rowNumberFromRef(value);
        }
        throw new Error("Unsupported ROW argument");
    }
    function evaluateEDate(args, context) {
        const startSerial = coerceDateSerial(evaluateFormulaAst(args[0], context));
        const months = Math.trunc(toNumber(evaluateFormulaAst(args[1], context)));
        const parts = excelSerialToDateParts(startSerial);
        const jsDate = new Date(Date.UTC(parts.year, parts.month - 1 + months, parts.day));
        return excelSerialFromDate(jsDate.getUTCFullYear(), jsDate.getUTCMonth() + 1, jsDate.getUTCDate());
    }
    function evaluateEoMonth(args, context) {
        const startSerial = coerceDateSerial(evaluateFormulaAst(args[0], context));
        const months = Math.trunc(toNumber(evaluateFormulaAst(args[1], context)));
        const parts = excelSerialToDateParts(startSerial);
        const jsDate = new Date(Date.UTC(parts.year, parts.month + months, 0));
        return excelSerialFromDate(jsDate.getUTCFullYear(), jsDate.getUTCMonth() + 1, jsDate.getUTCDate());
    }
    function evaluateCountIf(args, context) {
        const values = flattenValues(evaluateFormulaAst(args[0], context));
        const criteria = toText(evaluateFormulaAst(args[1], context));
        return values.filter((value) => matchesCriteria(value, criteria)).length;
    }
    function evaluateCount(args, context) {
        return args
            .flatMap((arg) => flattenValues(evaluateFormulaAst(arg, context)))
            .filter((value) => isCountableNumber(value))
            .length;
    }
    function evaluateCountA(args, context) {
        return args
            .flatMap((arg) => flattenValues(evaluateFormulaAst(arg, context)))
            .filter((value) => value !== null && value !== undefined && String(value) !== "")
            .length;
    }
    function evaluateSumIf(args, context) {
        var _a;
        const criteriaValues = flattenValues(evaluateFormulaAst(args[0], context));
        const criteria = toText(evaluateFormulaAst(args[1], context));
        const sumValues = args[2]
            ? flattenValues(evaluateFormulaAst(args[2], context))
            : criteriaValues;
        let total = 0;
        for (let index = 0; index < criteriaValues.length; index += 1) {
            if (matchesCriteria(criteriaValues[index], criteria)) {
                total += toNumber((_a = sumValues[index]) !== null && _a !== void 0 ? _a : 0);
            }
        }
        return total;
    }
    function evaluateCountIfs(args, context) {
        const criteriaPairs = [];
        for (let index = 0; index + 1 < args.length; index += 2) {
            criteriaPairs.push({
                values: flattenValues(evaluateFormulaAst(args[index], context)),
                criteria: toText(evaluateFormulaAst(args[index + 1], context))
            });
        }
        const maxLength = criteriaPairs.reduce((max, pair) => Math.max(max, pair.values.length), 0);
        let count = 0;
        for (let index = 0; index < maxLength; index += 1) {
            const matched = criteriaPairs.every((pair) => matchesCriteria(pair.values[index], pair.criteria));
            if (matched) {
                count += 1;
            }
        }
        return count;
    }
    function evaluateSumIfs(args, context) {
        var _a;
        const sumValues = flattenValues(evaluateFormulaAst(args[0], context));
        const criteriaPairs = [];
        for (let index = 1; index + 1 < args.length; index += 2) {
            criteriaPairs.push({
                values: flattenValues(evaluateFormulaAst(args[index], context)),
                criteria: toText(evaluateFormulaAst(args[index + 1], context))
            });
        }
        let total = 0;
        for (let index = 0; index < sumValues.length; index += 1) {
            const matched = criteriaPairs.every((pair) => matchesCriteria(pair.values[index], pair.criteria));
            if (matched) {
                total += toNumber((_a = sumValues[index]) !== null && _a !== void 0 ? _a : 0);
            }
        }
        return total;
    }
    function evaluateAverageIf(args, context) {
        var _a;
        const criteriaValues = flattenValues(evaluateFormulaAst(args[0], context));
        const criteria = toText(evaluateFormulaAst(args[1], context));
        const averageValues = args[2]
            ? flattenValues(evaluateFormulaAst(args[2], context))
            : criteriaValues;
        let total = 0;
        let count = 0;
        for (let index = 0; index < criteriaValues.length; index += 1) {
            if (matchesCriteria(criteriaValues[index], criteria)) {
                total += toNumber((_a = averageValues[index]) !== null && _a !== void 0 ? _a : 0);
                count += 1;
            }
        }
        return count === 0 ? "#DIV/0!" : total / count;
    }
    function evaluateAverageIfs(args, context) {
        var _a;
        const averageValues = flattenValues(evaluateFormulaAst(args[0], context));
        const criteriaPairs = [];
        for (let index = 1; index + 1 < args.length; index += 2) {
            criteriaPairs.push({
                values: flattenValues(evaluateFormulaAst(args[index], context)),
                criteria: toText(evaluateFormulaAst(args[index + 1], context))
            });
        }
        let total = 0;
        let count = 0;
        for (let index = 0; index < averageValues.length; index += 1) {
            const matched = criteriaPairs.every((pair) => matchesCriteria(pair.values[index], pair.criteria));
            if (matched) {
                total += toNumber((_a = averageValues[index]) !== null && _a !== void 0 ? _a : 0);
                count += 1;
            }
        }
        return count === 0 ? "#DIV/0!" : total / count;
    }
    function excelSerialFromDate(year, month, day) {
        const utcDate = Date.UTC(year, month - 1, day);
        const excelEpoch = Date.UTC(1899, 11, 30);
        return Math.floor((utcDate - excelEpoch) / 86400000);
    }
    function excelSerialToDateParts(serial) {
        const excelEpoch = Date.UTC(1899, 11, 30);
        const utcDate = new Date(excelEpoch + Math.floor(serial) * 86400000);
        return {
            year: utcDate.getUTCFullYear(),
            month: utcDate.getUTCMonth() + 1,
            day: utcDate.getUTCDate()
        };
    }
    function parseDateLikeString(value) {
        const normalized = value.replace(/[年\/.-]/g, "/").replace(/月/g, "/").replace(/日/g, "");
        const match = normalized.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
        if (!match) {
            return null;
        }
        return excelSerialFromDate(Number(match[1]), Number(match[2]), Number(match[3]));
    }
    function coerceDateSerial(value) {
        if (typeof value === "number") {
            return value;
        }
        const text = toText(value).trim();
        const parsed = parseDateLikeString(text);
        if (parsed !== null) {
            return parsed;
        }
        return toNumber(value);
    }
    function columnNumberFromRef(ref) {
        const match = String(ref).match(/\$?([A-Za-z]{1,3})\$?\d+/);
        if (!match) {
            throw new Error(`Invalid cell reference for COLUMN: ${ref}`);
        }
        const letters = match[1].toUpperCase();
        let number = 0;
        for (const char of letters) {
            number = number * 26 + (char.charCodeAt(0) - 64);
        }
        return number;
    }
    function rowNumberFromRef(ref) {
        const match = String(ref).match(/\$?[A-Za-z]{1,3}\$?(\d+)/);
        if (!match) {
            throw new Error(`Invalid cell reference for ROW: ${ref}`);
        }
        return Number(match[1]);
    }
    function toBoolean(value) {
        if (typeof value === "boolean") {
            return value;
        }
        if (typeof value === "number") {
            return value !== 0;
        }
        if (typeof value === "string") {
            if (!value) {
                return false;
            }
            const upper = value.toUpperCase();
            if (upper === "TRUE") {
                return true;
            }
            if (upper === "FALSE") {
                return false;
            }
            return true;
        }
        if (Array.isArray(value)) {
            return value.length > 0;
        }
        return Boolean(value);
    }
    function toNumber(value) {
        if (typeof value === "number") {
            return value;
        }
        if (typeof value === "boolean") {
            return value ? 1 : 0;
        }
        if (typeof value === "string") {
            const normalized = value.replace(/,/g, "");
            const parsed = Number(normalized);
            if (!Number.isNaN(parsed)) {
                return parsed;
            }
            if (value.toUpperCase() === "TRUE") {
                return 1;
            }
            if (value.toUpperCase() === "FALSE") {
                return 0;
            }
        }
        if (Array.isArray(value)) {
            return value.length;
        }
        return 0;
    }
    function toText(value) {
        if (value === null || value === undefined) {
            return "";
        }
        if (Array.isArray(value)) {
            return value.map((item) => toText(item)).join(":");
        }
        return String(value);
    }
    function flattenValues(value) {
        if (!Array.isArray(value)) {
            return [value];
        }
        if (value.length > 0 && Array.isArray(value[0])) {
            return value.flat();
        }
        return value;
    }
    function normalizeToMatrix(value) {
        if (!Array.isArray(value)) {
            return [[value]];
        }
        if (value.length > 0 && Array.isArray(value[0])) {
            return value;
        }
        return [value];
    }
    function matchesCriteria(value, criteria) {
        const trimmedCriteria = criteria.trim();
        const match = trimmedCriteria.match(/^(<=|>=|<>|=|<|>)(.*)$/);
        if (!match) {
            return looselyEquals(value, trimmedCriteria);
        }
        const operator = match[1];
        const rightRaw = match[2].trim();
        const leftNumeric = toNumber(value);
        const rightNumeric = Number(rightRaw.replace(/,/g, ""));
        if (!Number.isNaN(leftNumeric) && !Number.isNaN(rightNumeric)) {
            switch (operator) {
                case "<":
                    return leftNumeric < rightNumeric;
                case "<=":
                    return leftNumeric <= rightNumeric;
                case ">":
                    return leftNumeric > rightNumeric;
                case ">=":
                    return leftNumeric >= rightNumeric;
                case "=":
                    return leftNumeric === rightNumeric;
                case "<>":
                    return leftNumeric !== rightNumeric;
            }
        }
        const leftText = toText(value);
        switch (operator) {
            case "=":
                return leftText === rightRaw;
            case "<>":
                return leftText !== rightRaw;
            case "<":
                return leftText < rightRaw;
            case "<=":
                return leftText <= rightRaw;
            case ">":
                return leftText > rightRaw;
            case ">=":
                return leftText >= rightRaw;
            default:
                return false;
        }
    }
    function isCountableNumber(value) {
        if (typeof value === "number") {
            return true;
        }
        if (typeof value === "string") {
            const trimmed = value.trim();
            if (!trimmed) {
                return false;
            }
            return !Number.isNaN(Number(trimmed.replace(/,/g, "")));
        }
        return false;
    }
    function isFormulaError(value) {
        return typeof value === "string" && /^#(?:N\/A|REF!|VALUE!|DIV\/0!|NAME\?|NUM!|NULL!)/.test(value);
    }
    function looselyEquals(left, right) {
        if (typeof left === "number" || typeof right === "number") {
            return toNumber(left) === toNumber(right);
        }
        if (typeof left === "boolean" || typeof right === "boolean") {
            return toBoolean(left) === toBoolean(right);
        }
        return toText(left) === toText(right);
    }
    function compareValues(left, right) {
        if (typeof left === "number" || typeof right === "number") {
            return toNumber(left) - toNumber(right);
        }
        const leftText = toText(left);
        const rightText = toText(right);
        if (leftText === rightText) {
            return 0;
        }
        return leftText < rightText ? -1 : 1;
    }
    api.evaluateFormulaAst = evaluateFormulaAst;
    moduleRegistry.registerModule("formulaRuntime", api);
})(globalThis);
