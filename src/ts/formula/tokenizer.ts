/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */

(function initXlsx2mdFormulaTokenizer(global: typeof globalThis) {
  const moduleRegistry = getXlsx2mdModuleRegistry();
  const api = moduleRegistry.getModule<Record<string, unknown>>("formulaRuntime") || {};

  type FormulaTokenType =
    | "number"
    | "string"
    | "boolean"
    | "error"
    | "identifier"
    | "cell"
    | "quoted_identifier"
    | "lparen"
    | "rparen"
    | "lbrace"
    | "rbrace"
    | "comma"
    | "semicolon"
    | "colon"
    | "bang"
    | "lbracket"
    | "rbracket"
    | "operator";

  interface FormulaToken {
    type: FormulaTokenType;
    value: string;
    start: number;
    end: number;
  }

  const CELL_REF_RE = /^\$?[A-Za-z]{1,3}\$?\d+$/;
  const IDENTIFIER_START_RE = /[\p{L}_\\$]/u;
  const IDENTIFIER_PART_RE = /[\p{L}\p{N}_.\\$?]/u;

  function tokenizeFormula(input: string): FormulaToken[] {
    const source = normalizeFormulaInput(input);
    const tokens: FormulaToken[] = [];
    let index = 0;

    while (index < source.length) {
      const char = source[index];

      if (/\s/.test(char)) {
        const whitespaceStart = index;
        while (index < source.length && /\s/.test(source[index])) {
          index += 1;
        }
        const previousToken = tokens[tokens.length - 1] ?? null;
        const nextChar = source[index] ?? "";
        if (shouldEmitIntersectionOperator(previousToken, nextChar)) {
          tokens.push({
            type: "operator",
            value: " ",
            start: whitespaceStart,
            end: index
          });
        }
        continue;
      }

      const start = index;

      if (char === "\"") {
        const parsed = readStringLiteral(source, index);
        tokens.push({
          type: "string",
          value: parsed.value,
          start,
          end: parsed.end
        });
        index = parsed.end;
        continue;
      }

      if (char === "'") {
        const parsed = readQuotedIdentifier(source, index);
        tokens.push({
          type: "quoted_identifier",
          value: parsed.value,
          start,
          end: parsed.end
        });
        index = parsed.end;
        continue;
      }

      if (char === "#") {
        if (shouldReadErrorLiteral(source, index)) {
          const parsed = readErrorLiteral(source, index);
          tokens.push({
            type: "error",
            value: parsed.value,
            start,
            end: parsed.end
          });
          index = parsed.end;
          continue;
        }
        tokens.push({
          type: "operator",
          value: "#",
          start,
          end: start + 1
        });
        index += 1;
        continue;
      }

      if (/[0-9.]/.test(char)) {
        const parsed = readNumberLiteral(source, index);
        if (parsed) {
          tokens.push({
            type: "number",
            value: parsed.value,
            start,
            end: parsed.end
          });
          index = parsed.end;
          continue;
        }
      }

      if ("(),;:{}![]".includes(char)) {
        tokens.push({
          type: punctuationTypeFor(char),
          value: char,
          start,
          end: start + 1
        });
        index += 1;
        continue;
      }

      const operator = readOperator(source, index);
      if (operator) {
        tokens.push({
          type: "operator",
          value: operator,
          start,
          end: start + operator.length
        });
        index += operator.length;
        continue;
      }

      if (isIdentifierStart(char)) {
        const parsed = readIdentifierLike(source, index);
        const upperValue = parsed.value.toUpperCase();
        tokens.push({
          type: upperValue === "TRUE" || upperValue === "FALSE"
            ? "boolean"
            : isCellReference(parsed.value)
              ? "cell"
              : "identifier",
          value: parsed.value,
          start,
          end: parsed.end
        });
        index = parsed.end;
        continue;
      }

      throw new Error(`Unexpected formula token at ${index}: ${char}`);
    }

    return tokens;
  }

  function normalizeFormulaInput(input: string): string {
    return input.startsWith("=") ? input.slice(1) : input;
  }

  function readStringLiteral(source: string, start: number) {
    let index = start + 1;
    let value = "";
    while (index < source.length) {
      const char = source[index];
      if (char === "\"") {
        if (source[index + 1] === "\"") {
          value += "\"";
          index += 2;
          continue;
        }
        return { value, end: index + 1 };
      }
      value += char;
      index += 1;
    }
    throw new Error(`Unterminated string literal at ${start}`);
  }

  function readQuotedIdentifier(source: string, start: number) {
    let index = start + 1;
    let value = "";
    while (index < source.length) {
      const char = source[index];
      if (char === "'") {
        if (source[index + 1] === "'") {
          value += "'";
          index += 2;
          continue;
        }
        return { value, end: index + 1 };
      }
      value += char;
      index += 1;
    }
    throw new Error(`Unterminated quoted identifier at ${start}`);
  }

  function readErrorLiteral(source: string, start: number) {
    let index = start + 1;
    while (index < source.length && /[A-Za-z0-9/!?#]/.test(source[index])) {
      index += 1;
    }
    return { value: source.slice(start, index), end: index };
  }

  function readNumberLiteral(source: string, start: number) {
    const slice = source.slice(start);
    const match = slice.match(/^(?:\d+\.\d*|\.\d+|\d+)(?:[Ee][+\-]?\d+)?/);
    if (!match) {
      return null;
    }
    return {
      value: match[0],
      end: start + match[0].length
    };
  }

  function punctuationTypeFor(char: string): FormulaTokenType {
    switch (char) {
      case "(":
        return "lparen";
      case ")":
        return "rparen";
      case "{":
        return "lbrace";
      case "}":
        return "rbrace";
      case ",":
        return "comma";
      case ";":
        return "semicolon";
      case ":":
        return "colon";
      case "!":
        return "bang";
      case "[":
        return "lbracket";
      case "]":
        return "rbracket";
      default:
        throw new Error(`Unknown punctuation: ${char}`);
    }
  }

  function readOperator(source: string, start: number) {
    const twoChar = source.slice(start, start + 2);
    if (twoChar === "<>" || twoChar === "<=" || twoChar === ">=") {
      return twoChar;
    }
    const oneChar = source[start];
    return "+-*/&=<>%#".includes(oneChar) ? oneChar : null;
  }

  function shouldReadErrorLiteral(source: string, start: number) {
    return /^#(?:N\/A|REF!|VALUE!|NULL!|NUM!|NAME\?|DIV\/0!|CALC!|SPILL!|GETTING_DATA)/i.test(source.slice(start));
  }

  function shouldEmitIntersectionOperator(previousToken: FormulaToken | null, nextChar: string) {
    if (!previousToken) {
      return false;
    }
    const leftTokenTypes = new Set([
      "cell",
      "identifier",
      "quoted_identifier",
      "rparen",
      "rbracket",
      "rbrace"
    ]);
    if (!leftTokenTypes.has(previousToken.type)) {
      return false;
    }
    return nextChar === "'" || nextChar === "(" || isIdentifierStart(nextChar);
  }

  function isIdentifierStart(char: string) {
    return IDENTIFIER_START_RE.test(char);
  }

  function isIdentifierPart(char: string) {
    return IDENTIFIER_PART_RE.test(char);
  }

  function readIdentifierLike(source: string, start: number) {
    let index = start;
    while (index < source.length && isIdentifierPart(source[index])) {
      index += 1;
    }
    return {
      value: source.slice(start, index),
      end: index
    };
  }

  function isCellReference(value: string) {
    return CELL_REF_RE.test(value);
  }

  api.tokenizeFormula = tokenizeFormula;
  api.normalizeFormulaInput = normalizeFormulaInput;
  api.isCellReference = isCellReference;

  moduleRegistry.registerModule("formulaRuntime", api);
})(globalThis);
