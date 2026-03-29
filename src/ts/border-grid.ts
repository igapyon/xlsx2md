/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */

(() => {
  const moduleRegistry = getXlsx2mdModuleRegistry();
  type BorderFlags = {
    top: boolean;
    bottom: boolean;
    left: boolean;
    right: boolean;
  };

  type CellLike = {
    borders: BorderFlags;
    outputValue?: string;
  };

  type EdgeStats = {
    nonEmptyCount: number;
    borderCount: number;
    rawBorderCount: number;
    topCount: number;
    bottomCount: number;
    maxTextLength: number;
  };

  function getCellAt(cellMap: Map<string, CellLike>, row: number, col: number): CellLike | undefined {
    return cellMap.get(`${row}:${col}`);
  }

  function hasNormalizedBorderOnSide(
    cellMap: Map<string, CellLike>,
    row: number,
    col: number,
    side: keyof BorderFlags
  ): boolean {
    const cell = getCellAt(cellMap, row, col);
    if (side === "top") {
      const above = getCellAt(cellMap, row - 1, col);
      return !!cell?.borders.top || !!above?.borders.bottom;
    }
    if (side === "bottom") {
      const below = getCellAt(cellMap, row + 1, col);
      return !!cell?.borders.bottom || !!below?.borders.top;
    }
    if (side === "left") {
      const left = getCellAt(cellMap, row, col - 1);
      return !!cell?.borders.left || !!left?.borders.right;
    }
    const right = getCellAt(cellMap, row, col + 1);
    return !!cell?.borders.right || !!right?.borders.left;
  }

  function hasAnyNormalizedBorder(cellMap: Map<string, CellLike>, row: number, col: number): boolean {
    return hasNormalizedBorderOnSide(cellMap, row, col, "top")
      || hasNormalizedBorderOnSide(cellMap, row, col, "bottom")
      || hasNormalizedBorderOnSide(cellMap, row, col, "left")
      || hasNormalizedBorderOnSide(cellMap, row, col, "right");
  }

  function collectTableEdgeStats(
    cellMap: Map<string, CellLike>,
    row: number,
    startCol: number,
    endCol: number
  ): EdgeStats {
    let nonEmptyCount = 0;
    let borderCount = 0;
    let rawBorderCount = 0;
    let topCount = 0;
    let bottomCount = 0;
    let maxTextLength = 0;

    for (let col = startCol; col <= endCol; col += 1) {
      const cell = getCellAt(cellMap, row, col);
      const text = String(cell?.outputValue || "").trim();
      if (text) {
        nonEmptyCount += 1;
        maxTextLength = Math.max(maxTextLength, text.length);
      }
      if (hasAnyNormalizedBorder(cellMap, row, col)) {
        borderCount += 1;
      }
      if (cell && (cell.borders.top || cell.borders.bottom || cell.borders.left || cell.borders.right)) {
        rawBorderCount += 1;
      }
      if (hasNormalizedBorderOnSide(cellMap, row, col, "top")) {
        topCount += 1;
      }
      if (hasNormalizedBorderOnSide(cellMap, row, col, "bottom")) {
        bottomCount += 1;
      }
    }

    return { nonEmptyCount, borderCount, rawBorderCount, topCount, bottomCount, maxTextLength };
  }

  function countNormalizedBorderedCells(
    cellMap: Map<string, CellLike>,
    startRow: number,
    startCol: number,
    endRow: number,
    endCol: number
  ): number {
    let count = 0;
    for (let row = startRow; row <= endRow; row += 1) {
      for (let col = startCol; col <= endCol; col += 1) {
        if (hasAnyNormalizedBorder(cellMap, row, col)) {
          count += 1;
        }
      }
    }
    return count;
  }

  const borderGridApi = {
    getCellAt,
    hasNormalizedBorderOnSide,
    hasAnyNormalizedBorder,
    collectTableEdgeStats,
    countNormalizedBorderedCells
  };

  moduleRegistry.registerModule("borderGrid", borderGridApi);
})();
