export const XLSX2MD_CORE_TS_ORDER = [
  "src/ts/module-registry.ts",
  "src/ts/module-registry-access.ts",
  "src/ts/runtime-env.ts",
  "src/ts/office-drawing.ts",
  "src/ts/zip-io.ts",
  "src/ts/border-grid.ts",
  "src/ts/markdown-normalize.ts",
  "src/ts/markdown-escape.ts",
  "src/ts/markdown-table-escape.ts",
  "src/ts/rich-text-parser.ts",
  "src/ts/rich-text-plain-formatter.ts",
  "src/ts/rich-text-github-formatter.ts",
  "src/ts/rich-text-renderer.ts",
  "src/ts/narrative-structure.ts",
  "src/ts/table-detector.ts",
  "src/ts/markdown-export.ts",
  "src/ts/sheet-markdown.ts",
  "src/ts/styles-parser.ts",
  "src/ts/shared-strings.ts",
  "src/ts/address-utils.ts",
  "src/ts/rels-parser.ts",
  "src/ts/worksheet-tables.ts",
  "src/ts/cell-format.ts",
  "src/ts/xml-utils.ts",
  "src/ts/sheet-assets.ts",
  "src/ts/worksheet-parser.ts",
  "src/ts/workbook-loader.ts",
  "src/ts/formula-reference-utils.ts",
  "src/ts/formula-engine.ts",
  "src/ts/formula-legacy.ts",
  "src/ts/formula-ast.ts",
  "src/ts/formula-resolver.ts",
  "src/ts/formula/tokenizer.ts",
  "src/ts/formula/parser.ts",
  "src/ts/formula/evaluator.ts",
  "src/ts/core.ts"
];

export const XLSX2MD_APP_TS_ORDER = [
  ...XLSX2MD_CORE_TS_ORDER,
  "src/ts/main.ts"
];

export const XLSX2MD_CORE_JS_ORDER = XLSX2MD_CORE_TS_ORDER.map((filePath) => (
  filePath.replace("/ts/", "/js/").replace(/\.ts$/, ".js")
));
