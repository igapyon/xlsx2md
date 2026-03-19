import fs from "node:fs";
import path from "node:path";
import { buildSingleHtmlFromSource } from "./lib/single-html.mjs";

const ROOT = process.cwd();

const TARGETS = [
  {
    id: "xlsx2md",
    srcHtml: "xlsx2md-src.html",
    outHtml: "xlsx2md.html",
    tsOrder: [
      "src/xlsx2md/ts/office-drawing.ts",
      "src/xlsx2md/ts/zip-io.ts",
      "src/xlsx2md/ts/border-grid.ts",
      "src/xlsx2md/ts/narrative-structure.ts",
      "src/xlsx2md/ts/table-detector.ts",
      "src/xlsx2md/ts/markdown-export.ts",
      "src/xlsx2md/ts/styles-parser.ts",
      "src/xlsx2md/ts/shared-strings.ts",
      "src/xlsx2md/ts/worksheet-tables.ts",
      "src/xlsx2md/ts/formula/tokenizer.ts",
      "src/xlsx2md/ts/formula/parser.ts",
      "src/xlsx2md/ts/formula/evaluator.ts",
      "src/xlsx2md/ts/core.ts",
      "src/xlsx2md/ts/main.ts"
    ]
  }
];

const tsModule = await loadTypeScriptModule();

for (const target of TARGETS) {
  transpileTypeScript(target, tsModule);
  const srcPath = path.resolve(ROOT, target.srcHtml);
  const outPath = path.resolve(ROOT, target.outHtml);
  const source = fs.readFileSync(srcPath, "utf8");
  const output = buildSingleHtmlFromSource(source, srcPath, ROOT);
  fs.mkdirSync(path.dirname(outPath), { recursive: true });
  fs.writeFileSync(outPath, output, "utf8");
  console.log(`[build:xlsx2md] generated ${target.outHtml}`);
}

async function loadTypeScriptModule() {
  try {
    const module = await import("typescript");
    return module.default || module;
  } catch (_error) {
    return null;
  }
}

function transpileTypeScript(target, tsModule) {
  for (const relTsPath of target.tsOrder || []) {
    const tsPath = path.resolve(ROOT, relTsPath);
    const jsPath = path.resolve(
      ROOT,
      relTsPath.replace("/ts/", "/js/").replace(/\.ts$/, ".js")
    );

    const source = fs.readFileSync(tsPath, "utf8");
    let outputText = source;
    if (tsModule) {
      const result = tsModule.transpileModule(source, {
        compilerOptions: {
          target: tsModule.ScriptTarget.ES2019,
          module: tsModule.ModuleKind.None,
          lib: ["ES2020", "DOM"],
          strict: false,
          skipLibCheck: true
        },
        reportDiagnostics: true,
        fileName: tsPath
      });

      if (result.diagnostics && result.diagnostics.length > 0) {
        const errors = result.diagnostics
          .filter((diagnostic) => diagnostic.category === tsModule.DiagnosticCategory.Error)
          .map((diagnostic) => tsModule.flattenDiagnosticMessageText(diagnostic.messageText, "\n"));
        if (errors.length > 0) {
          throw new Error(`TypeScript transpile error in ${relTsPath}:\n${errors.join("\n")}`);
        }
      }
      outputText = result.outputText;
    } else {
      console.warn(
        `[build:xlsx2md] typescript not found. copied ${relTsPath} -> ${relTsPath.replace("/ts/", "/js/").replace(/\.ts$/, ".js")}`
      );
    }

    fs.mkdirSync(path.dirname(jsPath), { recursive: true });
    fs.writeFileSync(jsPath, outputText, "utf8");
  }
}
