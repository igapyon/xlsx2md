// @vitest-environment node

import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { loadXlsx2mdNodeApi } from "../scripts/lib/xlsx2md-node-runtime.mjs";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

function toArrayBuffer(buffer) {
  return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
}

describe("xlsx2md node runtime", () => {
  it("loads the core api and converts a workbook without browser ui", async () => {
    const api = loadXlsx2mdNodeApi({
      rootDir: path.resolve(__dirname, "..")
    });
    expect(globalThis.__xlsx2mdModuleRegistry?.getModule("xlsx2md")).toBe(api);
    const fixturePath = path.resolve(__dirname, "./fixtures/xlsx2md-basic-sample01.xlsx");
    const fixtureBytes = readFileSync(fixturePath);

    const workbook = await api.parseWorkbook(toArrayBuffer(fixtureBytes), "xlsx2md-basic-sample01.xlsx");
    const files = api.convertWorkbookToMarkdownFiles(workbook, {
      treatFirstRowAsHeader: true,
      trimText: true,
      removeEmptyRows: true,
      removeEmptyColumns: true
    });
    const combined = api.createCombinedMarkdownExportFile(workbook, files);

    expect(workbook.sheets.length).toBeGreaterThan(0);
    expect(files.length).toBe(workbook.sheets.length);
    expect(combined.fileName).toBe("xlsx2md-basic-sample01.md");
    expect(combined.content).toContain("# ");
  });
});
