(() => {
    const moduleRegistry = getXlsx2mdModuleRegistry();
    function createFormulaLegacyApi(deps) {
        function tryResolveFormulaExpressionLegacy(normalized, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries) {
            const ifResult = tryResolveIfFunction(normalized, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
            if (ifResult != null)
                return ifResult;
            const ifErrorResult = tryResolveIfErrorFunction(normalized, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
            if (ifErrorResult != null)
                return ifErrorResult;
            const logicalResult = tryResolveLogicalFunction(normalized, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
            if (logicalResult != null)
                return logicalResult;
            const concatResult = tryResolveConcatenationExpression(normalized, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
            if (concatResult != null)
                return concatResult;
            const numericFunctionResult = tryResolveNumericFunction(normalized, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
            if (numericFunctionResult != null)
                return numericFunctionResult;
            const datePartFunctionResult = tryResolveDatePartFunction(normalized, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
            if (datePartFunctionResult != null)
                return datePartFunctionResult;
            const predicateFunctionResult = tryResolvePredicateFunction(normalized, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
            if (predicateFunctionResult != null)
                return predicateFunctionResult;
            const chooseFunctionResult = tryResolveChooseFunction(normalized, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
            if (chooseFunctionResult != null)
                return chooseFunctionResult;
            const textFunctionResult = tryResolveTextFunction(normalized, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
            if (textFunctionResult != null)
                return textFunctionResult;
            const lookupFunctionResult = tryResolveLookupFunction(normalized, currentSheetName, resolveCellValue);
            if (lookupFunctionResult != null)
                return lookupFunctionResult;
            const stringFunctionResult = tryResolveStringFunction(normalized, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
            if (stringFunctionResult != null)
                return stringFunctionResult;
            const conditionalAggregateResult = tryResolveConditionalAggregateFunction(normalized, currentSheetName, resolveCellValue);
            if (conditionalAggregateResult != null)
                return conditionalAggregateResult;
            const aggregateResult = tryResolveAggregateFunction(normalized, currentSheetName, resolveRangeValues, resolveRangeEntries);
            if (aggregateResult != null)
                return aggregateResult;
            const comparisonResult = tryResolveComparisonExpression(normalized, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
            if (comparisonResult != null)
                return comparisonResult;
            if (/:/.test(normalized)) {
                return null;
            }
            const replacedRefs = normalized.replace(/(?:'((?:[^']|'')+)'|([A-Za-z0-9_ ]+))!(\$?[A-Z]+\$?\d+)|(\$?[A-Z]+\$?\d+)/g, (_full, quotedSheet, plainSheet, qualifiedAddress, localAddress) => {
                const sheetName = qualifiedAddress
                    ? deps.normalizeFormulaSheetName(quotedSheet || plainSheet || currentSheetName)
                    : currentSheetName;
                const address = deps.normalizeFormulaAddress(qualifiedAddress || localAddress || "");
                const rawValue = resolveCellValue(sheetName, address);
                const numericValue = Number(rawValue);
                if (rawValue === "" || Number.isNaN(numericValue)) {
                    throw new Error("__FORMULA_UNRESOLVED__");
                }
                return String(numericValue);
            });
            const replaced = replaceNumericDefinedNames(replacedRefs, currentSheetName);
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
                if (evaluations.some((value) => value === false))
                    return "FALSE";
                if (evaluations.some((value) => value == null))
                    return null;
                return evaluations.every(Boolean) ? "TRUE" : "FALSE";
            }
            if (functionName === "OR") {
                if (evaluations.some((value) => value === true))
                    return "TRUE";
                if (evaluations.some((value) => value == null))
                    return null;
                return evaluations.some(Boolean) ? "TRUE" : "FALSE";
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
            return deps.cellFormat.formatTextFunctionValue(value, formatText);
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
                    if (matchMode == null || !["0", ""].includes(matchMode.trim()))
                        return null;
                }
                if (args.length >= 6) {
                    const searchMode = resolveScalarFormulaValue(args[5], currentSheetName, resolveCellValue);
                    if (searchMode == null || !["1", ""].includes(searchMode.trim()))
                        return null;
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
                    if (matchType == null || !["0", ""].includes(matchType.trim()))
                        return null;
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
                const start = deps.parseCellAddress(rangeRef.start);
                const end = deps.parseCellAddress(rangeRef.end);
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
                return resolveCellValue(rangeRef.sheetName, `${deps.colToLetters(targetCol)}${targetRow}`);
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
                    if (!(normalizedLookup === "FALSE" || normalizedLookup === "0" || normalizedLookup === ""))
                        return null;
                }
                const start = deps.parseCellAddress(rangeRef.start);
                const end = deps.parseCellAddress(rangeRef.end);
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
                    const keyValue = resolveCellValue(rangeRef.sheetName, `${deps.colToLetters(col)}${startRow}`);
                    if (keyValue === "")
                        continue;
                    if (keyValue === lookupValue || (!Number.isNaN(Number(keyValue)) && !Number.isNaN(Number(lookupValue)) && Number(keyValue) === Number(lookupValue))) {
                        return resolveCellValue(rangeRef.sheetName, `${deps.colToLetters(col)}${targetRow}`);
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
            if (lookupValue == null || !rangeRef || Number.isNaN(columnIndex) || columnIndex < 1)
                return null;
            if (args.length === 4) {
                const rangeLookup = resolveScalarFormulaValue(args[3], currentSheetName, resolveCellValue);
                if (rangeLookup == null)
                    return null;
                const normalizedLookup = rangeLookup.trim().toUpperCase();
                if (!(normalizedLookup === "FALSE" || normalizedLookup === "0" || normalizedLookup === ""))
                    return null;
            }
            const start = deps.parseCellAddress(rangeRef.start);
            const end = deps.parseCellAddress(rangeRef.end);
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
                const keyValue = resolveCellValue(rangeRef.sheetName, `${deps.colToLetters(startCol)}${row}`);
                if (keyValue === "")
                    continue;
                if (keyValue === lookupValue || (!Number.isNaN(Number(keyValue)) && !Number.isNaN(Number(lookupValue)) && Number(keyValue) === Number(lookupValue))) {
                    return resolveCellValue(rangeRef.sheetName, `${deps.colToLetters(targetCol)}${row}`);
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
            const parts = deps.cellFormat.parseDateLikeParts(value);
            if (!parts)
                return null;
            if (fnName === "YEAR")
                return String(Number(parts.yyyy));
            if (fnName === "MONTH")
                return String(Number(parts.mm));
            if (fnName === "DAY")
                return String(Number(parts.dd));
            const returnType = args.length === 2
                ? resolveNumericFormulaArgument(args[1], currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries)
                : 1;
            if (returnType == null)
                return null;
            const weekday = new Date(Date.UTC(Number(parts.yyyy), Number(parts.mm) - 1, Number(parts.dd))).getUTCDay();
            return Math.trunc(returnType) === 2
                ? String(weekday === 0 ? 7 : weekday)
                : String(weekday + 1);
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
                const simpleRef = deps.parseSimpleFormulaReference(`=${args[0].trim()}`, currentSheetName);
                if (simpleRef) {
                    const value = resolveCellValue(simpleRef.sheetName, simpleRef.address);
                    return value.trim() === "" ? "TRUE" : "FALSE";
                }
                const value = resolveScalarFormulaValue(args[0], currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
                return value == null || value.trim() === "" ? "TRUE" : "FALSE";
            }
            const value = resolveScalarFormulaValue(args[0], currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
            if (fnName === "ISERROR") {
                if (value == null)
                    return "TRUE";
                return /^#(?:[A-Z]+\/[A-Z]+|[A-Z]+[!?]?)/i.test(value.trim()) ? "TRUE" : "FALSE";
            }
            if (fnName === "ISNA") {
                if (/^\s*VLOOKUP\(/i.test(args[0]))
                    return value == null ? "TRUE" : "FALSE";
                if (value == null)
                    return "FALSE";
                return /^#N\/A$/i.test(value.trim()) ? "TRUE" : "FALSE";
            }
            if (value == null)
                return "FALSE";
            if (fnName === "ISNUMBER") {
                if (value.trim() === "")
                    return "FALSE";
                return !Number.isNaN(Number(value)) ? "TRUE" : "FALSE";
            }
            if (fnName === "ISTEXT") {
                const normalized = value.trim().toUpperCase();
                if (normalized === "" || normalized === "TRUE" || normalized === "FALSE")
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
                if (resolved == null)
                    return null;
                values.push(resolved);
            }
            return values.join("");
        }
        function evaluateFormulaCondition(expression, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries) {
            const logical = tryResolveLogicalFunction(expression.trim(), currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
            if (logical != null)
                return logical === "TRUE";
            const comparison = tryResolveComparisonExpression(expression, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
            if (comparison != null)
                return comparison === "TRUE";
            const scalar = resolveScalarFormulaValue(expression, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
            if (scalar == null)
                return null;
            const normalized = scalar.trim().toUpperCase();
            if (normalized === "TRUE")
                return true;
            if (normalized === "FALSE")
                return false;
            const numeric = Number(scalar);
            return Number.isNaN(numeric) ? scalar.trim() !== "" : numeric !== 0;
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
            if (parts.length === 0)
                return null;
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
        function replaceNumericDefinedNames(expression, currentSheetName) {
            var _a;
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
                if ((expression[i] || "") === "(") {
                    result += token;
                    continue;
                }
                const scalar = ((_a = deps.getDefinedNameScalarValue()) === null || _a === void 0 ? void 0 : _a(currentSheetName, token)) || null;
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
                    if (resolved == null)
                        return segment;
                    const numericValue = Number(resolved);
                    if (Number.isNaN(numericValue))
                        return segment;
                    changed = true;
                    return String(numericValue);
                });
            }
            return current;
        }
        function resolveScalarFormulaValue(expression, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries) {
            var _a, _b;
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
            const simpleRef = deps.parseSimpleFormulaReference(`=${trimmed}`, currentSheetName);
            if (simpleRef) {
                return resolveCellValue(simpleRef.sheetName, simpleRef.address);
            }
            const scopedDefinedNameRef = deps.parseSheetScopedDefinedNameReference(trimmed, currentSheetName);
            if (scopedDefinedNameRef) {
                const scopedValue = ((_a = deps.getDefinedNameScalarValue()) === null || _a === void 0 ? void 0 : _a(scopedDefinedNameRef.sheetName, scopedDefinedNameRef.name)) || null;
                if (scopedValue != null)
                    return scopedValue;
            }
            const definedNameValue = ((_b = deps.getDefinedNameScalarValue()) === null || _b === void 0 ? void 0 : _b(currentSheetName, trimmed)) || null;
            if (definedNameValue != null)
                return definedNameValue;
            if (/^(TRUE|FALSE)$/i.test(trimmed)) {
                return trimmed.toUpperCase();
            }
            return deps.tryResolveFormulaExpression(`=${trimmed}`, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
        }
        function tryResolveAggregateFunction(normalizedFormula, currentSheetName, resolveRangeValues, resolveRangeEntries) {
            if (!resolveRangeValues || !resolveRangeEntries)
                return null;
            const call = parseWholeFunctionCall(normalizedFormula, ["SUM", "AVERAGE", "MIN", "MAX", "COUNT", "COUNTA"]);
            if (!call)
                return null;
            const fnName = call.name;
            const args = splitFormulaArguments(call.argsText.trim());
            if (args.length === 0)
                return null;
            const resolvedArgs = args.map((arg) => resolveAggregateArgument(arg, currentSheetName, resolveRangeValues, resolveRangeEntries));
            if (resolvedArgs.some((entry) => entry == null))
                return null;
            const values = resolvedArgs.flatMap((entry) => (entry === null || entry === void 0 ? void 0 : entry.numericValues) || []);
            const valueCount = resolvedArgs.reduce((sum, entry) => sum + ((entry === null || entry === void 0 ? void 0 : entry.valueCount) || 0), 0);
            if ((fnName !== "COUNTA" && values.length === 0) || valueCount === 0)
                return null;
            if (fnName === "SUM")
                return String(values.reduce((sum, value) => sum + value, 0));
            if (fnName === "AVERAGE")
                return String(values.reduce((sum, value) => sum + value, 0) / values.length);
            if (fnName === "MIN")
                return String(Math.min(...values));
            if (fnName === "MAX")
                return String(Math.max(...values));
            if (fnName === "COUNT")
                return String(values.length);
            if (fnName === "COUNTA")
                return String(valueCount);
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
                    if (!rangeCriteriaPairs.every((entry) => matchesCountIfCriteria(entry.cells[i], entry.criteria)))
                        continue;
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
                    if (!rangeCriteriaPairs.every((entry) => matchesCountIfCriteria(entry.cells[i], entry.criteria)))
                        continue;
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
                : parseQualifiedRangeReference(args[2] || args[0], currentSheetName);
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
            if (fnName === "COUNTIF")
                return String(count);
            if (fnName === "SUMIF")
                return String(sum);
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
                    const start = deps.parseCellAddress(rangeRef.start);
                    if (!start.row || !start.col)
                        return null;
                    return String(fnName === "ROW" ? start.row : start.col);
                }
                const simpleRef = deps.parseSimpleFormulaReference(`=${args[0]}`, currentSheetName);
                if (!simpleRef)
                    return null;
                const parsed = deps.parseCellAddress(simpleRef.address);
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
                const parsed = deps.cellFormat.parseValueFunctionText(source);
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
                const serial = deps.cellFormat.datePartsToExcelSerial(Math.trunc(year), Math.trunc(month), Math.trunc(day));
                return serial == null ? null : String(serial);
            }
            if (fnName === "EOMONTH") {
                if (args.length !== 2)
                    return null;
                const startValue = resolveScalarFormulaValue(args[0], currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
                const monthOffset = resolveNumericFormulaArgument(args[1], currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
                if (startValue == null || monthOffset == null)
                    return null;
                const parts = deps.cellFormat.parseDateLikeParts(startValue);
                if (!parts)
                    return null;
                const baseYear = Number(parts.yyyy);
                const baseMonthIndex = Number(parts.mm) - 1 + Math.trunc(monthOffset);
                const targetYear = baseYear + Math.floor(baseMonthIndex / 12);
                const targetMonth = ((baseMonthIndex % 12) + 12) % 12 + 1;
                const serial = deps.cellFormat.datePartsToExcelSerial(targetYear, targetMonth + 1, 0);
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
            if (fnName === "ROUND")
                return String(Math.round(value * factor) / factor);
            if (fnName === "ROUNDUP") {
                const scaled = value * factor;
                return String((scaled >= 0 ? Math.ceil(scaled) : Math.floor(scaled)) / factor);
            }
            if (fnName === "ROUNDDOWN") {
                const scaled = value * factor;
                return String((scaled >= 0 ? Math.floor(scaled) : Math.ceil(scaled)) / factor);
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
                return fnName === "LEN" ? String(source.length) : source.trim().replace(/\s+/g, " ");
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
                if (args.length === 3)
                    return source.split(oldText).join(newText);
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
                        if (evaluatedCondition == null)
                            return null;
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
            if (current.trim())
                args.push(current.trim());
            return args;
        }
        function resolveAggregateArgument(argText, currentSheetName, resolveRangeValues, resolveRangeEntries) {
            const rangeRef = parseQualifiedRangeReference(argText, currentSheetName);
            if (rangeRef) {
                const rangeEntries = resolveRangeEntries(rangeRef.sheetName, `${rangeRef.start}:${rangeRef.end}`);
                return {
                    numericValues: rangeEntries.numericValues,
                    valueCount: rangeEntries.rawValues.filter((value) => String(value || "").trim() !== "").length
                };
            }
            const numericLiteral = Number(argText);
            if (!Number.isNaN(numericLiteral)) {
                return { numericValues: [numericLiteral], valueCount: 1 };
            }
            const cellRef = deps.parseSimpleFormulaReference(`=${argText}`, currentSheetName);
            if (!cellRef)
                return null;
            const values = resolveRangeValues(cellRef.sheetName, `${cellRef.address}:${cellRef.address}`);
            const entryCount = resolveRangeEntries(cellRef.sheetName, `${cellRef.address}:${cellRef.address}`).rawValues
                .filter((value) => String(value || "").trim() !== "").length;
            return { numericValues: values, valueCount: entryCount };
        }
        function resolveNumericFormulaArgument(expression, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries) {
            const scalar = resolveScalarFormulaValue(expression, currentSheetName, resolveCellValue, resolveRangeValues, resolveRangeEntries);
            if (scalar == null)
                return null;
            const numeric = Number(scalar);
            return Number.isNaN(numeric) ? null : numeric;
        }
        function collectRangeCells(rangeRef, resolveCellValue) {
            const start = deps.parseCellAddress(rangeRef.start);
            const end = deps.parseCellAddress(rangeRef.end);
            if (!start.row || !start.col || !end.row || !end.col)
                return [];
            const startRow = Math.min(start.row, end.row);
            const endRow = Math.max(start.row, end.row);
            const startCol = Math.min(start.col, end.col);
            const endCol = Math.max(start.col, end.col);
            const values = [];
            for (let row = startRow; row <= endRow; row += 1) {
                for (let col = startCol; col <= endCol; col += 1) {
                    values.push(resolveCellValue(rangeRef.sheetName, `${deps.colToLetters(col)}${row}`));
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
            if (operator === "=")
                return numericComparable ? leftNum === rightNum : value === operandText;
            if (operator === "<>")
                return numericComparable ? leftNum !== rightNum : value !== operandText;
            if (!numericComparable)
                return false;
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
            var _a, _b, _c;
            const qualifiedRangeMatch = argText.match(/^(?:'((?:[^']|'')+)'|([^'=][^!]*))!(\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+)$/i);
            const localRangeMatch = argText.match(/^(\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+)$/i);
            if (!qualifiedRangeMatch && !localRangeMatch) {
                const scopedDefinedName = deps.parseSheetScopedDefinedNameReference(String(argText || "").trim(), currentSheetName);
                if (scopedDefinedName) {
                    const scopedRange = ((_a = deps.getDefinedNameRangeRef()) === null || _a === void 0 ? void 0 : _a(scopedDefinedName.sheetName, scopedDefinedName.name)) || null;
                    if (scopedRange)
                        return scopedRange;
                }
                const structuredRange = ((_b = deps.getStructuredRangeRef()) === null || _b === void 0 ? void 0 : _b(currentSheetName, String(argText || "").trim())) || null;
                if (structuredRange)
                    return structuredRange;
                const definedRange = ((_c = deps.getDefinedNameRangeRef()) === null || _c === void 0 ? void 0 : _c(currentSheetName, String(argText || "").trim())) || null;
                if (definedRange)
                    return definedRange;
                return null;
            }
            const sheetName = qualifiedRangeMatch
                ? deps.normalizeFormulaSheetName(qualifiedRangeMatch[1] || qualifiedRangeMatch[2] || currentSheetName)
                : currentSheetName;
            const rangeText = String(qualifiedRangeMatch ? qualifiedRangeMatch[3] : (localRangeMatch === null || localRangeMatch === void 0 ? void 0 : localRangeMatch[1]) || "");
            const range = deps.parseRangeAddress(rangeText);
            if (!range)
                return null;
            return { sheetName, start: range.start, end: range.end };
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
                    if (tokens[index] !== ")")
                        throw new Error("Unbalanced parentheses");
                    index += 1;
                    return value;
                }
                if (token == null)
                    throw new Error("Unexpected end of expression");
                index += 1;
                const numericValue = Number(token);
                if (Number.isNaN(numericValue))
                    throw new Error("Invalid token");
                return numericValue;
            }
            const result = parseExpression();
            if (index !== tokens.length)
                throw new Error("Unexpected trailing tokens");
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
                if (!numberMatch)
                    throw new Error("Invalid arithmetic expression");
                tokens.push(numberMatch[0]);
                index += numberMatch[0].length;
            }
            return tokens;
        }
        return {
            tryResolveFormulaExpressionLegacy,
            findTopLevelOperatorIndex,
            parseWholeFunctionCall,
            splitFormulaArguments,
            parseQualifiedRangeReference,
            resolveScalarFormulaValue
        };
    }
    const formulaLegacyApi = {
        createFormulaLegacyApi
    };
    moduleRegistry.registerModule("formulaLegacy", formulaLegacyApi);
})();
