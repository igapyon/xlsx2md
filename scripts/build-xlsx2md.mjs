import fs from "node:fs";
import path from "node:path";
import { buildSingleHtmlFromSource } from "./lib/single-html.mjs";
import { XLSX2MD_APP_TS_ORDER } from "./lib/xlsx2md-module-order.mjs";

const ROOT = process.cwd();
const BUILD_DATE_PLACEHOLDER = "{{BUILD_DATE}}";

const TARGETS = [
  {
    id: "index",
    srcHtml: "index-src.html",
    outHtml: "index.html"
  },
  {
    id: "xlsx2md",
    srcHtml: "xlsx2md-src.html",
    outHtml: "xlsx2md.html",
    tsOrder: XLSX2MD_APP_TS_ORDER
  }
];

const tsModule = await loadTypeScriptModule();

for (const target of TARGETS) {
  transpileTypeScript(target, tsModule);
  const srcPath = path.resolve(ROOT, target.srcHtml);
  const outPath = path.resolve(ROOT, target.outHtml);
  const source = fs.readFileSync(srcPath, "utf8");
  const sourceWithBuildDate = source.replaceAll(BUILD_DATE_PLACEHOLDER, formatBuildDate());
  const output = buildSingleHtmlFromSource(sourceWithBuildDate, srcPath, ROOT);
  fs.mkdirSync(path.dirname(outPath), { recursive: true });
  fs.writeFileSync(outPath, output, "utf8");
  console.log(`[build:xlsx2md] generated ${target.outHtml}`);
}

function formatBuildDate(date = new Date()) {
  const yyyy = String(date.getFullYear());
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const dd = String(date.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
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
