export const XLSX2MD_CORE_TS_ORDER = [
  "src/xlsx2md/ts/module-registry.ts",
  "src/xlsx2md/ts/module-registry-access.ts",
  "src/xlsx2md/ts/runtime-env.ts",
  "src/xlsx2md/ts/office-drawing.ts",
  "src/xlsx2md/ts/zip-io.ts",
  "src/xlsx2md/ts/border-grid.ts",
  "src/xlsx2md/ts/markdown-normalize.ts",
  "src/xlsx2md/ts/markdown-escape.ts",
  "src/xlsx2md/ts/markdown-table-escape.ts",
  "src/xlsx2md/ts/rich-text-parser.ts",
  "src/xlsx2md/ts/rich-text-plain-formatter.ts",
  "src/xlsx2md/ts/rich-text-github-formatter.ts",
  "src/xlsx2md/ts/rich-text-renderer.ts",
  "src/xlsx2md/ts/narrative-structure.ts",
  "src/xlsx2md/ts/table-detector.ts",
  "src/xlsx2md/ts/markdown-export.ts",
  "src/xlsx2md/ts/sheet-markdown.ts",
  "src/xlsx2md/ts/styles-parser.ts",
  "src/xlsx2md/ts/shared-strings.ts",
  "src/xlsx2md/ts/address-utils.ts",
  "src/xlsx2md/ts/rels-parser.ts",
  "src/xlsx2md/ts/worksheet-tables.ts",
  "src/xlsx2md/ts/cell-format.ts",
  "src/xlsx2md/ts/xml-utils.ts",
  "src/xlsx2md/ts/sheet-assets.ts",
  "src/xlsx2md/ts/worksheet-parser.ts",
  "src/xlsx2md/ts/workbook-loader.ts",
  "src/xlsx2md/ts/formula-reference-utils.ts",
  "src/xlsx2md/ts/formula-engine.ts",
  "src/xlsx2md/ts/formula-legacy.ts",
  "src/xlsx2md/ts/formula-ast.ts",
  "src/xlsx2md/ts/formula-resolver.ts",
  "src/xlsx2md/ts/formula/tokenizer.ts",
  "src/xlsx2md/ts/formula/parser.ts",
  "src/xlsx2md/ts/formula/evaluator.ts",
  "src/xlsx2md/ts/core.ts"
];

export const XLSX2MD_APP_TS_ORDER = [
  ...XLSX2MD_CORE_TS_ORDER,
  "src/xlsx2md/ts/main.ts"
];

export const XLSX2MD_CORE_JS_ORDER = XLSX2MD_CORE_TS_ORDER.map((filePath) => (
  filePath.replace("/ts/", "/js/").replace(/\.ts$/, ".js")
));
