// @vitest-environment jsdom

import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";
import { loadModuleRegistry } from "./helpers/module-registry.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const borderGridCode = readFileSync(
  path.resolve(__dirname, "../src/js/border-grid.js"),
  "utf8"
);

function bootBorderGrid() {
  document.body.innerHTML = "";
  loadModuleRegistry(__dirname);
  new Function(borderGridCode)();
  return globalThis.__xlsx2mdModuleRegistry.getModule("borderGrid");
}

describe("xlsx2md border grid", () => {
  it("returns false when neither adjacent cell owns the queried edge", () => {
    const api = bootBorderGrid();
    const cellMap = new Map([
      ["1:1", { borders: { top: false, right: false, bottom: false, left: false }, outputValue: "A" }],
      ["2:1", { borders: { top: false, right: false, bottom: false, left: false }, outputValue: "B" }]
    ]);

    expect(api.hasNormalizedBorderOnSide(cellMap, 2, 1, "top")).toBe(false);
    expect(api.hasAnyNormalizedBorder(cellMap, 2, 1)).toBe(false);
  });

  it("normalizes a horizontal border from the cell above", () => {
    const api = bootBorderGrid();
    const cellMap = new Map([
      ["1:1", { borders: { top: false, right: false, bottom: true, left: false }, outputValue: "Header" }],
      ["2:1", { borders: { top: false, right: false, bottom: false, left: false }, outputValue: "Note" }]
    ]);

    expect(api.hasNormalizedBorderOnSide(cellMap, 2, 1, "top")).toBe(true);
    expect(api.hasAnyNormalizedBorder(cellMap, 2, 1)).toBe(true);
  });

  it("normalizes a vertical border from the cell on the left", () => {
    const api = bootBorderGrid();
    const cellMap = new Map([
      ["1:1", { borders: { top: false, right: true, bottom: false, left: false }, outputValue: "A" }],
      ["1:2", { borders: { top: false, right: false, bottom: false, left: false }, outputValue: "B" }]
    ]);

    expect(api.hasNormalizedBorderOnSide(cellMap, 1, 2, "left")).toBe(true);
    expect(api.hasAnyNormalizedBorder(cellMap, 1, 2)).toBe(true);
  });

  it("treats outer frame cells as bordered even without inner grid lines", () => {
    const api = bootBorderGrid();
    const cellMap = new Map([
      ["1:1", { borders: { top: true, right: false, bottom: false, left: true }, outputValue: "A1" }],
      ["1:2", { borders: { top: true, right: true, bottom: false, left: false }, outputValue: "B1" }],
      ["2:1", { borders: { top: false, right: false, bottom: true, left: true }, outputValue: "A2" }],
      ["2:2", { borders: { top: false, right: true, bottom: true, left: false }, outputValue: "B2" }]
    ]);

    expect(api.countNormalizedBorderedCells(cellMap, 1, 1, 2, 2)).toBe(4);
    expect(api.hasNormalizedBorderOnSide(cellMap, 1, 1, "left")).toBe(true);
    expect(api.hasNormalizedBorderOnSide(cellMap, 1, 2, "right")).toBe(true);
    expect(api.hasNormalizedBorderOnSide(cellMap, 2, 1, "bottom")).toBe(true);
  });

  it("normalizes an internal horizontal separator to both touching rows", () => {
    const api = bootBorderGrid();
    const cellMap = new Map([
      ["1:1", { borders: { top: false, right: false, bottom: true, left: false }, outputValue: "H1" }],
      ["1:2", { borders: { top: false, right: false, bottom: true, left: false }, outputValue: "H2" }],
      ["2:1", { borders: { top: false, right: false, bottom: false, left: false }, outputValue: "D1" }],
      ["2:2", { borders: { top: false, right: false, bottom: false, left: false }, outputValue: "D2" }]
    ]);

    const firstRow = api.collectTableEdgeStats(cellMap, 1, 1, 2);
    const secondRow = api.collectTableEdgeStats(cellMap, 2, 1, 2);

    expect(firstRow.bottomCount).toBe(2);
    expect(secondRow.topCount).toBe(2);
    expect(secondRow.rawBorderCount).toBe(0);
  });

  it("normalizes an internal vertical separator to both touching columns", () => {
    const api = bootBorderGrid();
    const cellMap = new Map([
      ["1:1", { borders: { top: false, right: true, bottom: false, left: false }, outputValue: "L" }],
      ["1:2", { borders: { top: false, right: false, bottom: false, left: false }, outputValue: "R" }]
    ]);

    expect(api.hasNormalizedBorderOnSide(cellMap, 1, 1, "right")).toBe(true);
    expect(api.hasNormalizedBorderOnSide(cellMap, 1, 2, "left")).toBe(true);
  });

  it("collects row stats with normalized and raw border counts separately", () => {
    const api = bootBorderGrid();
    const cellMap = new Map([
      ["1:1", { borders: { top: true, right: true, bottom: true, left: true }, outputValue: "H1" }],
      ["1:2", { borders: { top: true, right: true, bottom: true, left: true }, outputValue: "H2" }],
      ["2:1", { borders: { top: false, right: false, bottom: false, left: false }, outputValue: "Note line" }],
      ["2:2", { borders: { top: false, right: false, bottom: false, left: false }, outputValue: "" }]
    ]);

    const firstRow = api.collectTableEdgeStats(cellMap, 1, 1, 2);
    const secondRow = api.collectTableEdgeStats(cellMap, 2, 1, 2);

    expect(firstRow.borderCount).toBe(2);
    expect(firstRow.rawBorderCount).toBe(2);
    expect(firstRow.bottomCount).toBe(2);

    expect(secondRow.nonEmptyCount).toBe(1);
    expect(secondRow.rawBorderCount).toBe(0);
    expect(secondRow.topCount).toBe(2);
    expect(secondRow.borderCount).toBe(2);
    expect(secondRow.maxTextLength).toBe("Note line".length);
  });

  it("counts only non-empty cells for text stats while still seeing borders on empty cells", () => {
    const api = bootBorderGrid();
    const cellMap = new Map([
      ["3:1", { borders: { top: true, right: false, bottom: false, left: false }, outputValue: "" }],
      ["3:2", { borders: { top: true, right: false, bottom: false, left: false }, outputValue: "Caption" }],
      ["3:3", { borders: { top: true, right: false, bottom: false, left: false }, outputValue: "" }]
    ]);

    const rowStats = api.collectTableEdgeStats(cellMap, 3, 1, 3);

    expect(rowStats.nonEmptyCount).toBe(1);
    expect(rowStats.borderCount).toBe(3);
    expect(rowStats.rawBorderCount).toBe(3);
    expect(rowStats.topCount).toBe(3);
    expect(rowStats.maxTextLength).toBe("Caption".length);
  });

  it("counts normalized bordered cells across a candidate range", () => {
    const api = bootBorderGrid();
    const cellMap = new Map([
      ["1:1", { borders: { top: true, right: false, bottom: false, left: true }, outputValue: "A1" }],
      ["1:2", { borders: { top: true, right: true, bottom: false, left: false }, outputValue: "B1" }],
      ["2:1", { borders: { top: false, right: false, bottom: true, left: true }, outputValue: "A2" }],
      ["2:2", { borders: { top: false, right: true, bottom: true, left: false }, outputValue: "B2" }],
      ["3:1", { borders: { top: false, right: false, bottom: false, left: false }, outputValue: "Tail" }]
    ]);

    expect(api.countNormalizedBorderedCells(cellMap, 1, 1, 2, 2)).toBe(4);
    expect(api.countNormalizedBorderedCells(cellMap, 3, 1, 3, 1)).toBe(1);
  });

  it("does not count cells outside the requested range", () => {
    const api = bootBorderGrid();
    const cellMap = new Map([
      ["1:1", { borders: { top: true, right: true, bottom: true, left: true }, outputValue: "A1" }],
      ["1:2", { borders: { top: true, right: true, bottom: true, left: true }, outputValue: "B1" }],
      ["2:1", { borders: { top: true, right: true, bottom: true, left: true }, outputValue: "A2" }],
      ["2:2", { borders: { top: true, right: true, bottom: true, left: true }, outputValue: "B2" }]
    ]);

    expect(api.countNormalizedBorderedCells(cellMap, 1, 1, 1, 1)).toBe(1);
    expect(api.countNormalizedBorderedCells(cellMap, 1, 2, 2, 2)).toBe(2);
  });
});
