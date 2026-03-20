import fs from "node:fs/promises";
import path from "node:path";

import { loadXlsx2mdNodeApi } from "./lib/xlsx2md-node-runtime.mjs";

function printHelp() {
  console.log(`Usage:
  node scripts/xlsx2md-cli.mjs <input.xlsx> [options]

Options:
  --out <file>                  Write combined Markdown to this file
  --zip <file>                  Write ZIP export to this file
  --output-mode <mode>          display | raw | both (default: display)
  --include-shape-details       Include shape source details
  --no-header-row               Do not treat first row as table header
  --no-trim-text                Preserve surrounding whitespace
  --keep-empty-rows             Keep empty rows
  --keep-empty-columns          Keep empty columns
  --summary                     Print per-sheet summary to stdout
  --help                        Show this help
`);
}

function parseArgs(argv) {
  const options = {
    treatFirstRowAsHeader: true,
    trimText: true,
    removeEmptyRows: true,
    removeEmptyColumns: true,
    includeShapeDetails: false,
    outputMode: "display",
    summary: false,
    outPath: null,
    zipPath: null
  };
  const positionals = [];

  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    if (!arg.startsWith("--")) {
      positionals.push(arg);
      continue;
    }

    if (arg === "--help") {
      options.help = true;
      continue;
    }
    if (arg === "--include-shape-details") {
      options.includeShapeDetails = true;
      continue;
    }
    if (arg === "--no-header-row") {
      options.treatFirstRowAsHeader = false;
      continue;
    }
    if (arg === "--no-trim-text") {
      options.trimText = false;
      continue;
    }
    if (arg === "--keep-empty-rows") {
      options.removeEmptyRows = false;
      continue;
    }
    if (arg === "--keep-empty-columns") {
      options.removeEmptyColumns = false;
      continue;
    }
    if (arg === "--summary") {
      options.summary = true;
      continue;
    }
    if (arg === "--out" || arg === "--zip" || arg === "--output-mode") {
      const value = argv[index + 1];
      if (!value) {
        throw new Error(`Missing value for ${arg}`);
      }
      index += 1;
      if (arg === "--out") options.outPath = value;
      if (arg === "--zip") options.zipPath = value;
      if (arg === "--output-mode") {
        if (value !== "display" && value !== "raw" && value !== "both") {
          throw new Error(`Invalid output mode: ${value}`);
        }
        options.outputMode = value;
      }
      continue;
    }

    throw new Error(`Unknown option: ${arg}`);
  }

  if (positionals.length > 1) {
    throw new Error("Only one input workbook can be specified.");
  }
  options.inputPath = positionals[0] || null;
  return options;
}

function toArrayBuffer(buffer) {
  return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
}

async function main() {
  const options = parseArgs(process.argv.slice(2));
  if (options.help || !options.inputPath) {
    printHelp();
    process.exit(options.help ? 0 : 1);
  }

  const inputPath = path.resolve(options.inputPath);
  const inputBytes = await fs.readFile(inputPath);
  const api = loadXlsx2mdNodeApi();
  const workbook = await api.parseWorkbook(toArrayBuffer(inputBytes), path.basename(inputPath));
  const files = api.convertWorkbookToMarkdownFiles(workbook, {
    treatFirstRowAsHeader: options.treatFirstRowAsHeader,
    trimText: options.trimText,
    removeEmptyRows: options.removeEmptyRows,
    removeEmptyColumns: options.removeEmptyColumns,
    includeShapeDetails: options.includeShapeDetails,
    outputMode: options.outputMode
  });

  if (options.summary) {
    for (const file of files) {
      console.log(api.createSummaryText(file));
      console.log("");
    }
  }

  if (options.zipPath) {
    const zipBytes = api.createWorkbookExportArchive(workbook, files);
    const zipPath = path.resolve(options.zipPath);
    await fs.mkdir(path.dirname(zipPath), { recursive: true });
    await fs.writeFile(zipPath, zipBytes);
  }

  if (!options.zipPath || options.outPath) {
    const combined = api.createCombinedMarkdownExportFile(workbook, files);
    const outputPath = path.resolve(options.outPath || combined.fileName);
    await fs.mkdir(path.dirname(outputPath), { recursive: true });
    await fs.writeFile(outputPath, `${combined.content}\n`, "utf8");
  }
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exit(1);
});
