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
  --include-shape-details       Include shape source details in Markdown
  --no-header-row               Do not treat the first row as a table header
  --no-trim-text                Preserve surrounding whitespace
  --keep-empty-rows             Keep empty rows
  --keep-empty-columns          Keep empty columns
  --summary                     Print per-sheet summary to stdout
  --help                        Show this help and exit

Exit codes:
  0                             Success
  1                             Error
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

  if (positionals.length === 0) {
    options.inputPath = null;
  } else if (positionals.length === 1) {
    [options.inputPath] = positionals;
  } else {
    throw new Error("Specify exactly one input workbook.");
  }
  return options;
}

function toArrayBuffer(buffer) {
  return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
}

async function writeTextFile(outputPath, content) {
  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  await fs.writeFile(outputPath, content, "utf8");
}

async function writeBinaryFile(outputPath, content) {
  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  await fs.writeFile(outputPath, content);
}

function formatWorkbookError(inputPath, stage, error) {
  const message = error instanceof Error ? error.message : String(error);
  return `[${path.basename(inputPath)}] ${stage}: ${message}`;
}

function printWorkbookSummary(api, workbookName, files) {
  console.log(`[workbook] ${workbookName}`);
  for (const file of files) {
    console.log(api.createSummaryText(file));
    console.log("");
  }
}

async function main() {
  const options = parseArgs(process.argv.slice(2));
  if (options.help || !options.inputPath) {
    printHelp();
    process.exit(options.help ? 0 : 1);
  }

  const api = loadXlsx2mdNodeApi();
  const inputPath = path.resolve(options.inputPath);

  try {
    let inputBytes;
    try {
      inputBytes = await fs.readFile(inputPath);
    } catch (error) {
      throw new Error(formatWorkbookError(inputPath, "read failed", error));
    }

    let workbook;
    try {
      workbook = await api.parseWorkbook(toArrayBuffer(inputBytes), path.basename(inputPath));
    } catch (error) {
      throw new Error(formatWorkbookError(inputPath, "parse failed", error));
    }

    let files;
    try {
      files = api.convertWorkbookToMarkdownFiles(workbook, {
        treatFirstRowAsHeader: options.treatFirstRowAsHeader,
        trimText: options.trimText,
        removeEmptyRows: options.removeEmptyRows,
        removeEmptyColumns: options.removeEmptyColumns,
        includeShapeDetails: options.includeShapeDetails,
        outputMode: options.outputMode
      });
    } catch (error) {
      throw new Error(formatWorkbookError(inputPath, "convert failed", error));
    }

    if (options.summary) {
      printWorkbookSummary(api, path.basename(inputPath), files);
    }

    const combined = api.createCombinedMarkdownExportFile(workbook, files);

    if (options.zipPath) {
      try {
        const zipBytes = api.createWorkbookExportArchive(workbook, files);
        await writeBinaryFile(path.resolve(options.zipPath), zipBytes);
      } catch (error) {
        throw new Error(formatWorkbookError(inputPath, "zip write failed", error));
      }
    }

    if (!options.zipPath || options.outPath) {
      const markdownOutputPath = options.outPath
        ? path.resolve(options.outPath)
        : path.resolve(combined.fileName);
      try {
        await writeTextFile(markdownOutputPath, `${combined.content}\n`);
      } catch (error) {
        throw new Error(formatWorkbookError(inputPath, "markdown write failed", error));
      }
    }
  } catch (error) {
    if (error instanceof Error && error.message.startsWith(`[${path.basename(inputPath)}] `)) {
      throw error;
    }
    throw new Error(formatWorkbookError(inputPath, "failed", error));
  }
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exit(1);
});
