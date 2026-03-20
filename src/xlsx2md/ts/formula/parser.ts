(function initXlsx2mdFormulaParser(global: typeof globalThis) {
  const moduleRegistry = getXlsx2mdModuleRegistry();
  const api = moduleRegistry.getModule<Record<string, unknown>>("formulaRuntime");
  if (!api) {
    throw new Error("xlsx2md formula runtime module is not loaded");
  }

  type FormulaAstNode =
    | { type: "number"; value: number; raw: string; }
    | { type: "string"; value: string; }
    | { type: "boolean"; value: boolean; raw: string; }
    | { type: "error"; value: string; }
    | { type: "array_constant"; rows: FormulaAstNode[][]; }
    | { type: "name"; name: string; }
    | { type: "scoped_name"; sheet: string; name: string; }
    | { type: "cell"; ref: string; sheet: string | null; }
    | { type: "range"; start: FormulaAstNode; end: FormulaAstNode; }
    | { type: "structured_ref"; table: string; column: string; qualifier?: string | null; }
    | { type: "function_call"; name: string; args: FormulaAstNode[]; }
    | { type: "unary_op"; operator: string; operand: FormulaAstNode; }
    | { type: "postfix_op"; operator: string; operand: FormulaAstNode; }
    | { type: "binary_op"; operator: string; left: FormulaAstNode; right: FormulaAstNode; };

  interface FormulaToken {
    type: string;
    value: string;
    start: number;
    end: number;
  }

  interface ParserState {
    tokens: FormulaToken[];
    index: number;
  }

  function parseFormula(input: string): FormulaAstNode {
    const tokens = api.tokenizeFormula(input);
    const state: ParserState = { tokens, index: 0 };
    const ast = parseComparison(state);
    if (peek(state)) {
      throw new Error(`Unexpected trailing token: ${peek(state)?.value}`);
    }
    return ast;
  }

  function parseComparison(state: ParserState): FormulaAstNode {
    let left = parseConcat(state);
    while (matchOperator(state, ["=", "<>", "<", "<=", ">", ">="])) {
      const operator = consume(state).value;
      const right = parseConcat(state);
      left = { type: "binary_op", operator, left, right };
    }
    return left;
  }

  function parseConcat(state: ParserState): FormulaAstNode {
    let left = parseAdditive(state);
    while (matchOperator(state, ["&"])) {
      const operator = consume(state).value;
      const right = parseAdditive(state);
      left = { type: "binary_op", operator, left, right };
    }
    return left;
  }

  function parseAdditive(state: ParserState): FormulaAstNode {
    let left = parseMultiplicative(state);
    while (matchOperator(state, ["+", "-"])) {
      const operator = consume(state).value;
      const right = parseMultiplicative(state);
      left = { type: "binary_op", operator, left, right };
    }
    return left;
  }

  function parseMultiplicative(state: ParserState): FormulaAstNode {
    let left = parseIntersection(state);
    while (matchOperator(state, ["*", "/"])) {
      const operator = consume(state).value;
      const right = parseIntersection(state);
      left = { type: "binary_op", operator, left, right };
    }
    return left;
  }

  function parseIntersection(state: ParserState): FormulaAstNode {
    let left = parseUnary(state);
    while (matchOperator(state, [" "])) {
      const operator = consume(state).value;
      const right = parseUnary(state);
      left = { type: "binary_op", operator, left, right };
    }
    return left;
  }

  function parseUnary(state: ParserState): FormulaAstNode {
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

  function parsePostfix(state: ParserState): FormulaAstNode {
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

  function parsePrimary(state: ParserState): FormulaAstNode {
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

  function parseReferenceLike(state: ParserState): FormulaAstNode {
    const first = consume(state);

    if (first.type === "identifier" && peek(state)?.type === "lparen") {
      return parseFunctionCall(state, first.value);
    }

    if ((first.type === "identifier" || first.type === "quoted_identifier") && peek(state)?.type === "lbracket") {
      return parseStructuredReference(state, first.value);
    }

    if (peek(state)?.type === "bang") {
      consume(state);
      const next = consume(state);
      if (!next || (next.type !== "cell" && next.type !== "identifier")) {
        throw new Error(`Expected reference after !, got ${next?.value ?? "EOF"}`);
      }
      let node: FormulaAstNode = next.type === "cell"
        ? { type: "cell", ref: next.value, sheet: first.value }
        : { type: "scoped_name", sheet: first.value, name: next.value };
      if (peek(state)?.type === "colon") {
        consume(state);
        const end = parseRangeEndpoint(state, first.value);
        node = { type: "range", start: node, end };
      }
      return node;
    }

    if (first.type === "cell") {
      const cellNode: FormulaAstNode = { type: "cell", ref: first.value, sheet: null };
      if (peek(state)?.type === "colon") {
        consume(state);
        const end = parseRangeEndpoint(state, null);
        return { type: "range", start: cellNode, end };
      }
      return cellNode;
    }

    return { type: "name", name: first.value };
  }

  function parseStructuredReference(state: ParserState, tableName: string): FormulaAstNode {
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

  function readStructuredReferenceSegment(state: ParserState): string {
    let text = "";
    while (peek(state) && peek(state)?.type !== "rbracket") {
      const token = consume(state);
      if (!token || !["identifier", "quoted_identifier", "cell", "error", "number", "boolean", "operator"].includes(token.type)) {
        throw new Error(`Expected structured reference column, got ${token?.value ?? "EOF"}`);
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

  function parseFunctionCall(state: ParserState, name: string): FormulaAstNode {
    expect(state, "lparen");
    const args: FormulaAstNode[] = [];
    if (peek(state)?.type !== "rparen") {
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

  function parseArrayConstant(state: ParserState): FormulaAstNode {
    expect(state, "lbrace");
    const rows: FormulaAstNode[][] = [];
    if (peek(state)?.type !== "rbrace") {
      while (true) {
        const row: FormulaAstNode[] = [];
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

  function parseRangeEndpoint(state: ParserState, defaultSheet: string | null): FormulaAstNode {
    const token = consume(state);
    if (!token || (token.type !== "cell" && token.type !== "identifier")) {
      throw new Error(`Expected range endpoint, got ${token?.value ?? "EOF"}`);
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

  function peek(state: ParserState) {
    return state.tokens[state.index] ?? null;
  }

  function consume(state: ParserState) {
    const token = state.tokens[state.index] ?? null;
    if (token) {
      state.index += 1;
    }
    return token;
  }

  function expect(state: ParserState, type: string) {
    const token = consume(state);
    if (!token || token.type !== type) {
      throw new Error(`Expected ${type}, got ${token?.type ?? "EOF"}`);
    }
    return token;
  }

  function matchOperator(state: ParserState, operators: string[]) {
    const token = peek(state);
    return token?.type === "operator" && operators.includes(token.value);
  }

  function matchAndConsume(state: ParserState, type: string) {
    if (peek(state)?.type === type) {
      consume(state);
      return true;
    }
    return false;
  }

  api.parseFormula = parseFormula;

  moduleRegistry.registerModule("formulaRuntime", api);
})(globalThis);
