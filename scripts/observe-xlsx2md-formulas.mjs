import { Blob as NodeBlob } from "node:buffer";
import { readdirSync, readFileSync } from "node:fs";
import path from "node:path";
import { DecompressionStream as NodeDecompressionStream } from "node:stream/web";

import { JSDOM } from "jsdom";

const ROOT = process.cwd();
const LOCAL_DATA_DIR = path.resolve(ROOT, "local-data");
const tokenizerCode = readFileSync(path.resolve(ROOT, "src/js/formula/tokenizer.js"), "utf8");
const parserCode = readFileSync(path.resolve(ROOT, "src/js/formula/parser.js"), "utf8");
const evaluatorCode = readFileSync(path.resolve(ROOT, "src/js/formula/evaluator.js"), "utf8");
const coreCode = readFileSync(path.resolve(ROOT, "src/js/core.js"), "utf8");

const dom = new JSDOM("<!doctype html><html><body></body></html>");
globalThis.window = dom.window;
globalThis.document = dom.window.document;
globalThis.DOMParser = dom.window.DOMParser;
globalThis.Node = dom.window.Node;
globalThis.Blob = NodeBlob;
globalThis.DecompressionStream = NodeDecompressionStream;

new Function(tokenizerCode)();
new Function(parserCode)();
new Function(evaluatorCode)();
new Function(coreCode)();

const formulaApi = globalThis.__xlsx2mdFormula;
const xlsxApi = globalThis.__xlsx2md;

const workbookFiles = readdirSync(LOCAL_DATA_DIR)
  .filter((name) => /\.(xlsx|xlsm)$/i.test(name))
  .sort();

for (const fileName of workbookFiles) {
  const filePath = path.resolve(LOCAL_DATA_DIR, fileName);
  const fileBytes = readFileSync(filePath);
  const arrayBuffer = fileBytes.buffer.slice(fileBytes.byteOffset, fileBytes.byteOffset + fileBytes.byteLength);
  const workbook = await xlsxApi.parseWorkbook(arrayBuffer, fileName);

  const workbookSummary = [];
  const functionCounts = new Map();

  for (const sheet of workbook.sheets) {
    const formulaCells = sheet.cells.filter((cell) => cell.formulaText);
    let astParseOk = 0;
    let astParseNg = 0;

    for (const cell of formulaCells) {
      try {
        const ast = formulaApi.parseFormula(cell.formulaText);
        astParseOk += 1;
        collectFunctionNames(ast, functionCounts);
      } catch (_error) {
        astParseNg += 1;
      }
    }

    const fallbackCount = formulaCells.filter((cell) => cell.resolutionStatus === "fallback_formula").length;
    workbookSummary.push({
      sheet: sheet.name,
      formulas: formulaCells.length,
      fallback: fallbackCount,
      astParseOk,
      astParseNg
    });
  }

  console.log(`# ${fileName}`);
  for (const item of workbookSummary) {
    console.log(
      `- ${item.sheet}: formulas=${item.formulas}, fallback=${item.fallback}, ast_ok=${item.astParseOk}, ast_ng=${item.astParseNg}`
    );
  }
  const topFunctions = Array.from(functionCounts.entries())
    .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))
    .slice(0, 12);
  console.log(`- top_functions: ${topFunctions.map(([name, count]) => `${name}:${count}`).join(", ")}`);
  console.log("");
}

function collectFunctionNames(ast, functionCounts) {
  if (!ast || typeof ast !== "object") {
    return;
  }
  if (ast.type === "function_call") {
    const upperName = String(ast.name || "").toUpperCase();
    functionCounts.set(upperName, (functionCounts.get(upperName) || 0) + 1);
    for (const arg of ast.args || []) {
      collectFunctionNames(arg, functionCounts);
    }
    return;
  }
  if (ast.type === "binary_op") {
    collectFunctionNames(ast.left, functionCounts);
    collectFunctionNames(ast.right, functionCounts);
    return;
  }
  if (ast.type === "unary_op" || ast.type === "postfix_op") {
    collectFunctionNames(ast.operand, functionCounts);
    return;
  }
  if (ast.type === "range") {
    collectFunctionNames(ast.start, functionCounts);
    collectFunctionNames(ast.end, functionCounts);
  }
}
