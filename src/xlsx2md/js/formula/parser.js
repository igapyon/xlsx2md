(function initXlsx2mdFormulaParser(global) {
    var _a;
    var _b;
    const api = ((_a = (_b = global).__xlsx2mdFormula) !== null && _a !== void 0 ? _a : (_b.__xlsx2mdFormula = {}));
    function parseFormula(input) {
        var _a;
        const tokens = api.tokenizeFormula(input);
        const state = { tokens, index: 0 };
        const ast = parseComparison(state);
        if (peek(state)) {
            throw new Error(`Unexpected trailing token: ${(_a = peek(state)) === null || _a === void 0 ? void 0 : _a.value}`);
        }
        return ast;
    }
    function parseComparison(state) {
        let left = parseConcat(state);
        while (matchOperator(state, ["=", "<>", "<", "<=", ">", ">="])) {
            const operator = consume(state).value;
            const right = parseConcat(state);
            left = { type: "binary_op", operator, left, right };
        }
        return left;
    }
    function parseConcat(state) {
        let left = parseAdditive(state);
        while (matchOperator(state, ["&"])) {
            const operator = consume(state).value;
            const right = parseAdditive(state);
            left = { type: "binary_op", operator, left, right };
        }
        return left;
    }
    function parseAdditive(state) {
        let left = parseMultiplicative(state);
        while (matchOperator(state, ["+", "-"])) {
            const operator = consume(state).value;
            const right = parseMultiplicative(state);
            left = { type: "binary_op", operator, left, right };
        }
        return left;
    }
    function parseMultiplicative(state) {
        let left = parseIntersection(state);
        while (matchOperator(state, ["*", "/"])) {
            const operator = consume(state).value;
            const right = parseIntersection(state);
            left = { type: "binary_op", operator, left, right };
        }
        return left;
    }
    function parseIntersection(state) {
        let left = parseUnary(state);
        while (matchOperator(state, [" "])) {
            const operator = consume(state).value;
            const right = parseUnary(state);
            left = { type: "binary_op", operator, left, right };
        }
        return left;
    }
    function parseUnary(state) {
        if (matchOperator(state, ["+", "-"])) {
            const operator = consume(state).value;
            return {
                type: "unary_op",
                operator,
                operand: parseUnary(state)
            };
        }
        return parsePostfix(state);
    }
    function parsePostfix(state) {
        let node = parsePrimary(state);
        while (matchOperator(state, ["%", "#"])) {
            const operator = consume(state).value;
            node = {
                type: "postfix_op",
                operator,
                operand: node
            };
        }
        return node;
    }
    function parsePrimary(state) {
        const token = peek(state);
        if (!token) {
            throw new Error("Unexpected end of formula");
        }
        if (token.type === "number") {
            consume(state);
            return {
                type: "number",
                value: Number(token.value),
                raw: token.value
            };
        }
        if (token.type === "string") {
            consume(state);
            return {
                type: "string",
                value: token.value
            };
        }
        if (token.type === "boolean") {
            consume(state);
            return {
                type: "boolean",
                value: token.value.toUpperCase() === "TRUE",
                raw: token.value
            };
        }
        if (token.type === "error") {
            consume(state);
            return {
                type: "error",
                value: token.value
            };
        }
        if (token.type === "lbrace") {
            return parseArrayConstant(state);
        }
        if (token.type === "lparen") {
            consume(state);
            const expression = parseComparison(state);
            expect(state, "rparen");
            return expression;
        }
        if (token.type === "identifier" || token.type === "cell" || token.type === "quoted_identifier") {
            return parseReferenceLike(state);
        }
        throw new Error(`Unexpected token in formula: ${token.value}`);
    }
    function parseReferenceLike(state) {
        var _a, _b, _c, _d, _e, _f;
        const first = consume(state);
        if (first.type === "identifier" && ((_a = peek(state)) === null || _a === void 0 ? void 0 : _a.type) === "lparen") {
            return parseFunctionCall(state, first.value);
        }
        if ((first.type === "identifier" || first.type === "quoted_identifier") && ((_b = peek(state)) === null || _b === void 0 ? void 0 : _b.type) === "lbracket") {
            return parseStructuredReference(state, first.value);
        }
        if (((_c = peek(state)) === null || _c === void 0 ? void 0 : _c.type) === "bang") {
            consume(state);
            const next = consume(state);
            if (!next || (next.type !== "cell" && next.type !== "identifier")) {
                throw new Error(`Expected reference after !, got ${(_d = next === null || next === void 0 ? void 0 : next.value) !== null && _d !== void 0 ? _d : "EOF"}`);
            }
            let node = next.type === "cell"
                ? { type: "cell", ref: next.value, sheet: first.value }
                : { type: "scoped_name", sheet: first.value, name: next.value };
            if (((_e = peek(state)) === null || _e === void 0 ? void 0 : _e.type) === "colon") {
                consume(state);
                const end = parseRangeEndpoint(state, first.value);
                node = { type: "range", start: node, end };
            }
            return node;
        }
        if (first.type === "cell") {
            const cellNode = { type: "cell", ref: first.value, sheet: null };
            if (((_f = peek(state)) === null || _f === void 0 ? void 0 : _f.type) === "colon") {
                consume(state);
                const end = parseRangeEndpoint(state, null);
                return { type: "range", start: cellNode, end };
            }
            return cellNode;
        }
        return { type: "name", name: first.value };
    }
    function parseStructuredReference(state, tableName) {
        expect(state, "lbracket");
        if (matchAndConsume(state, "lbracket")) {
            const qualifier = readStructuredReferenceSegment(state);
            expect(state, "rbracket");
            expect(state, "comma");
            expect(state, "lbracket");
            const column = readStructuredReferenceSegment(state);
            expect(state, "rbracket");
            expect(state, "rbracket");
            return {
                type: "structured_ref",
                table: tableName,
                qualifier,
                column
            };
        }
        const column = readStructuredReferenceSegment(state);
        expect(state, "rbracket");
        return {
            type: "structured_ref",
            table: tableName,
            column
        };
    }
    function readStructuredReferenceSegment(state) {
        var _a, _b;
        let text = "";
        while (peek(state) && ((_a = peek(state)) === null || _a === void 0 ? void 0 : _a.type) !== "rbracket") {
            const token = consume(state);
            if (!token || !["identifier", "quoted_identifier", "cell", "error", "number", "boolean", "operator"].includes(token.type)) {
                throw new Error(`Expected structured reference column, got ${(_b = token === null || token === void 0 ? void 0 : token.value) !== null && _b !== void 0 ? _b : "EOF"}`);
            }
            if (token.type === "operator" && token.value !== "#" && token.value !== " ") {
                throw new Error(`Expected structured reference column, got ${token.value}`);
            }
            text += token.value;
        }
        if (!text.length) {
            throw new Error("Expected structured reference column, got EOF");
        }
        return text.startsWith("#")
            ? `#${text.slice(1).replace(/\s+/g, " ").trim()}`
            : text;
    }
    function parseFunctionCall(state, name) {
        var _a;
        expect(state, "lparen");
        const args = [];
        if (((_a = peek(state)) === null || _a === void 0 ? void 0 : _a.type) !== "rparen") {
            do {
                args.push(parseComparison(state));
            } while (matchAndConsume(state, "comma"));
        }
        expect(state, "rparen");
        return {
            type: "function_call",
            name,
            args
        };
    }
    function parseArrayConstant(state) {
        var _a;
        expect(state, "lbrace");
        const rows = [];
        if (((_a = peek(state)) === null || _a === void 0 ? void 0 : _a.type) !== "rbrace") {
            while (true) {
                const row = [];
                row.push(parseComparison(state));
                while (matchAndConsume(state, "comma")) {
                    row.push(parseComparison(state));
                }
                rows.push(row);
                if (!matchAndConsume(state, "semicolon")) {
                    break;
                }
            }
        }
        expect(state, "rbrace");
        return {
            type: "array_constant",
            rows
        };
    }
    function parseRangeEndpoint(state, defaultSheet) {
        var _a;
        const token = consume(state);
        if (!token || (token.type !== "cell" && token.type !== "identifier")) {
            throw new Error(`Expected range endpoint, got ${(_a = token === null || token === void 0 ? void 0 : token.value) !== null && _a !== void 0 ? _a : "EOF"}`);
        }
        if (token.type === "cell") {
            return {
                type: "cell",
                ref: token.value,
                sheet: defaultSheet
            };
        }
        return {
            type: defaultSheet ? "scoped_name" : "name",
            ...(defaultSheet
                ? { sheet: defaultSheet, name: token.value }
                : { name: token.value })
        };
    }
    function peek(state) {
        var _a;
        return (_a = state.tokens[state.index]) !== null && _a !== void 0 ? _a : null;
    }
    function consume(state) {
        var _a;
        const token = (_a = state.tokens[state.index]) !== null && _a !== void 0 ? _a : null;
        if (token) {
            state.index += 1;
        }
        return token;
    }
    function expect(state, type) {
        var _a;
        const token = consume(state);
        if (!token || token.type !== type) {
            throw new Error(`Expected ${type}, got ${(_a = token === null || token === void 0 ? void 0 : token.type) !== null && _a !== void 0 ? _a : "EOF"}`);
        }
        return token;
    }
    function matchOperator(state, operators) {
        const token = peek(state);
        return (token === null || token === void 0 ? void 0 : token.type) === "operator" && operators.includes(token.value);
    }
    function matchAndConsume(state, type) {
        var _a;
        if (((_a = peek(state)) === null || _a === void 0 ? void 0 : _a.type) === type) {
            consume(state);
            return true;
        }
        return false;
    }
    api.parseFormula = parseFormula;
})(globalThis);
