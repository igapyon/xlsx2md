// @vitest-environment jsdom

import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const borderGridCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/border-grid.js"),
  "utf8"
);
const tableDetectorCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/table-detector.js"),
  "utf8"
);

function bootTableDetector() {
  document.body.innerHTML = "";
  new Function(borderGridCode)();
  new Function(tableDetectorCode)();
  return globalThis.__xlsx2mdTableDetector;
}

function createCell(row, col, outputValue, borders = {}) {
  return {
    row,
    col,
    outputValue,
    borders: {
      top: false,
      bottom: false,
      left: false,
      right: false,
      ...borders
    }
  };
}

function buildCellMap(sheet) {
  const cellMap = new Map();
  for (const cell of sheet.cells) {
    cellMap.set(`${cell.row}:${cell.col}`, cell);
  }
  return cellMap;
}

describe("xlsx2md table detector", () => {
  it("collects seed cells from values or borders", () => {
    const api = bootTableDetector();
    const sheet = {
      cells: [
        createCell(1, 1, ""),
        createCell(1, 2, "項目"),
        createCell(2, 1, "", { bottom: true })
      ],
      merges: []
    };

    const seeds = api.collectTableSeedCells(sheet);

    expect(seeds.map((cell) => `${cell.row}:${cell.col}`)).toEqual(["1:2", "2:1"]);
  });

  it("breaks a bordered table before a following borderless note row", () => {
    const api = bootTableDetector();
    const cellMap = new Map([
      ["1:1", createCell(1, 1, "項番", { top: true, bottom: true, left: true })],
      ["1:2", createCell(1, 2, "名称", { top: true, bottom: true, right: true })],
      ["2:1", createCell(2, 1, "1", { bottom: true, left: true })],
      ["2:2", createCell(2, 2, "コード", { bottom: true, right: true })],
      ["3:1", createCell(3, 1, "※注記")],
      ["3:2", createCell(3, 2, "")]
    ]);

    const trimmed = api.trimTableCandidateBounds(cellMap, {
      startRow: 1,
      startCol: 1,
      endRow: 3,
      endCol: 2
    });

    expect(trimmed).toEqual({
      startRow: 1,
      startCol: 1,
      endRow: 2,
      endCol: 2
    });
  });

  it("normalizes candidate matrices by trimming empty rows and columns while keeping merge markers non-meaningful", () => {
    const api = bootTableDetector();
    const sheet = {
      cells: [
        createCell(1, 1, "Header"),
        createCell(1, 2, ""),
        createCell(1, 3, ""),
        createCell(2, 1, "Value"),
        createCell(2, 2, "A"),
        createCell(2, 3, ""),
        createCell(3, 1, ""),
        createCell(3, 2, ""),
        createCell(3, 3, "")
      ],
      merges: [
        { startRow: 1, startCol: 1, endRow: 1, endCol: 2, ref: "A1:B1" }
      ]
    };

    const matrix = api.matrixFromCandidate(
      sheet,
      { startRow: 1, startCol: 1, endRow: 3, endCol: 3, score: 5, reasonSummary: [] },
      { trimText: true, removeEmptyRows: true, removeEmptyColumns: true },
      buildCellMap,
      (cell) => cell?.outputValue || ""
    );

    expect(matrix).toEqual([
      ["Header", "[MERGED←]"],
      ["Value", "A"]
    ]);
  });

  it("detects a bordered dense grid as a table candidate", () => {
    const api = bootTableDetector();
    const sheet = {
      cells: [
        createCell(1, 1, "項番", { top: true, bottom: true, left: true }),
        createCell(1, 2, "名称", { top: true, bottom: true, right: true }),
        createCell(2, 1, "1", { bottom: true, left: true }),
        createCell(2, 2, "コード", { bottom: true, right: true })
      ],
      merges: []
    };

    const candidates = api.detectTableCandidates(sheet, buildCellMap);

    expect(candidates).toHaveLength(1);
    expect(candidates[0]).toMatchObject({
      startRow: 1,
      startCol: 1,
      endRow: 2,
      endCol: 2
    });
    expect(candidates[0].score).toBeGreaterThanOrEqual(api.defaultTableScoreWeights.threshold);
  });
});
