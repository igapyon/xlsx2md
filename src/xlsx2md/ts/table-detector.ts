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
    row: number;
    col: number;
    outputValue: string;
    borders: BorderFlags;
  };

  type MergeRange = {
    startRow: number;
    startCol: number;
    endRow: number;
    endCol: number;
    ref: string;
  };

  type SheetLike<TCell extends CellLike = CellLike> = {
    cells: TCell[];
    merges: MergeRange[];
  };

  type TableCandidate = {
    startRow: number;
    startCol: number;
    endRow: number;
    endCol: number;
    score: number;
    reasonSummary: string[];
  };

  type MarkdownOptions = {
    trimText?: boolean;
    removeEmptyRows?: boolean;
    removeEmptyColumns?: boolean;
  };

  type TableDetectionMode = "balanced" | "border";

  type TableScoreWeights = {
    minGrid: number;
    borderPresence: number;
    densityHigh: number;
    densityVeryHigh: number;
    headerish: number;
    mergeHeavyPenalty: number;
    prosePenalty: number;
    threshold: number;
  };

  const DEFAULT_TABLE_SCORE_WEIGHTS: TableScoreWeights = {
    minGrid: 2,
    borderPresence: 3,
    densityHigh: 2,
    densityVeryHigh: 1,
    headerish: 2,
    mergeHeavyPenalty: -1,
    prosePenalty: -2,
    threshold: 4
  };

  const borderGridHelper = moduleRegistry?.getModule<{
    collectTableEdgeStats: <T extends { borders: BorderFlags; outputValue?: string }>(cellMap: Map<string, T>, row: number, startCol: number, endCol: number) => {
      nonEmptyCount: number;
      borderCount: number;
      rawBorderCount: number;
      topCount: number;
      bottomCount: number;
      maxTextLength: number;
      };
      countNormalizedBorderedCells: <T extends { borders: BorderFlags }>(cellMap: Map<string, T>, startRow: number, startCol: number, endRow: number, endCol: number) => number;
  }>("borderGrid");
  if (!borderGridHelper) {
    throw new Error("xlsx2md border grid module is not loaded");
  }

  function collectTableSeedCells<TCell extends CellLike>(sheet: SheetLike<TCell>): TCell[] {
    return sheet.cells.filter((cell) => {
      const hasValue = !!String(cell.outputValue || "").trim();
      const hasBorder = cell.borders.top || cell.borders.bottom || cell.borders.left || cell.borders.right;
      return hasValue || hasBorder;
    });
  }

  function collectBorderSeedCells<TCell extends CellLike>(sheet: SheetLike<TCell>): TCell[] {
    return sheet.cells.filter((cell) => (
      cell.borders.top || cell.borders.bottom || cell.borders.left || cell.borders.right
    ));
  }

  function areBorderAdjacent<TCell extends CellLike>(current: TCell, next: TCell): boolean {
    if (current.row === next.row && Math.abs(current.col - next.col) === 1) {
      return (current.borders.top && next.borders.top)
        || (current.borders.bottom && next.borders.bottom)
        || (current.col < next.col ? current.borders.right && next.borders.left : current.borders.left && next.borders.right);
    }
    if (current.col === next.col && Math.abs(current.row - next.row) === 1) {
      return (current.borders.left && next.borders.left)
        || (current.borders.right && next.borders.right)
        || (current.row < next.row ? current.borders.bottom && next.borders.top : current.borders.top && next.borders.bottom);
    }
    return false;
  }

  function collectConnectedComponents<TCell extends CellLike>(seedCells: TCell[], adjacencyMode: "grid" | "border" = "grid"): TCell[][] {
    const positionMap = new Map<string, TCell>();
    for (const cell of seedCells) {
      positionMap.set(`${cell.row}:${cell.col}`, cell);
    }
    const visited = new Set<string>();
    const components: TCell[][] = [];

    for (const cell of seedCells) {
      const key = `${cell.row}:${cell.col}`;
      if (visited.has(key)) continue;
      const queue = [cell];
      const component: TCell[] = [];
      visited.add(key);
      while (queue.length > 0) {
        const current = queue.shift() as TCell;
        component.push(current);
        for (const [rowDelta, colDelta] of [[1, 0], [-1, 0], [0, 1], [0, -1]]) {
          const nextKey = `${current.row + rowDelta}:${current.col + colDelta}`;
          const nextCell = positionMap.get(nextKey);
          if (!nextCell || visited.has(nextKey)) continue;
          if (adjacencyMode === "border" && !areBorderAdjacent(current, nextCell)) continue;
          visited.add(nextKey);
          queue.push(nextCell);
        }
      }
      components.push(component);
    }

    return components;
  }

  function isWithinBounds(
    bounds: { startRow: number; startCol: number; endRow: number; endCol: number },
    candidate: { startRow: number; startCol: number; endRow: number; endCol: number }
  ): boolean {
    return candidate.startRow >= bounds.startRow
      && candidate.startCol >= bounds.startCol
      && candidate.endRow <= bounds.endRow
      && candidate.endCol <= bounds.endCol;
  }

  function getBoundsArea(bounds: { startRow: number; startCol: number; endRow: number; endCol: number }): number {
    return Math.max(1, (bounds.endRow - bounds.startRow + 1) * (bounds.endCol - bounds.startCol + 1));
  }

  function getCombinedCandidateArea(
    candidates: Array<{ startRow: number; startCol: number; endRow: number; endCol: number }>
  ): number {
    return candidates.reduce((sum, candidate) => sum + getBoundsArea(candidate), 0);
  }

  function pruneRedundantCandidates(candidates: TableCandidate[]): TableCandidate[] {
    return candidates.filter((candidate, candidateIndex) => {
      const candidateArea = getBoundsArea(candidate);
      const hasSingleDominatingContainedCandidate = candidates.some((other, otherIndex) => {
        if (candidateIndex === otherIndex) return false;
        if (!isWithinBounds(candidate, other)) return false;
        const otherArea = getBoundsArea(other);
        if (otherArea < candidateArea * 0.4) return false;
        return candidateArea > otherArea;
      });
      if (hasSingleDominatingContainedCandidate) {
        return false;
      }
      const containedCandidates = candidates.filter((other, otherIndex) => {
        if (candidateIndex === otherIndex) return false;
        if (!isWithinBounds(candidate, other)) return false;
        return getBoundsArea(other) < candidateArea;
      });
      if (containedCandidates.length >= 2 && getCombinedCandidateArea(containedCandidates) >= candidateArea * 0.6) {
        return false;
      }
      return true;
    });
  }

  function detectTableCandidates<TCell extends CellLike>(
    sheet: SheetLike<TCell>,
    buildCellMap: (sheet: SheetLike<TCell>) => Map<string, TCell>,
    scoreWeights: TableScoreWeights = DEFAULT_TABLE_SCORE_WEIGHTS,
    tableDetectionMode: TableDetectionMode = "balanced"
  ): TableCandidate[] {
    const cellMap = buildCellMap(sheet);
    const allSeedCells = collectTableSeedCells(sheet);
    const borderSeedCells = collectBorderSeedCells(sheet);
    const candidates: TableCandidate[] = [];
    const candidateKeys = new Set<string>();

    function maybePushCandidate(component: TCell[], sourceKind: "border" | "fallback" = "border"): void {
      const rows = component.map((entry) => entry.row);
      const cols = component.map((entry) => entry.col);
      const startRow = Math.min(...rows);
      const endRow = Math.max(...rows);
      const startCol = Math.min(...cols);
      const endCol = Math.max(...cols);
      const area = Math.max(1, (endRow - startRow + 1) * (endCol - startCol + 1));
      const density = component.filter((entry) => entry.outputValue.trim()).length / area;
      const rowCount = endRow - startRow + 1;
      const colCount = endCol - startCol + 1;
      if (rowCount < 2 || colCount < 2) {
        return;
      }

      let score = 0;
      const reasons: string[] = [];
      const normalizedBorderedCellCount = borderGridHelper.countNormalizedBorderedCells(cellMap, startRow, startCol, endRow, endCol);
      if (rowCount >= 2 && colCount >= 2) {
        score += scoreWeights.minGrid;
        reasons.push(`At least 2x2 (+${scoreWeights.minGrid})`);
      }
      if (normalizedBorderedCellCount >= Math.max(2, Math.ceil(component.length * 0.3))) {
        score += scoreWeights.borderPresence;
        reasons.push(`Has borders (+${scoreWeights.borderPresence})`);
      }
      if (density >= 0.55) {
        score += scoreWeights.densityHigh;
        reasons.push(`High density (+${scoreWeights.densityHigh})`);
      }
      if (density >= 0.8) {
        score += scoreWeights.densityVeryHigh;
        reasons.push(`Very high density (+${scoreWeights.densityVeryHigh})`);
      }

      const firstRowCells = component.filter((entry) => entry.row === startRow).sort((a, b) => a.col - b.col);
      const headerishCount = firstRowCells.filter((entry) => {
        const value = entry.outputValue.trim();
        return value.length > 0 && value.length <= 24 && !/^\d+(?:\.\d+)?$/.test(value);
      }).length;
      if (headerishCount >= 2) {
        score += scoreWeights.headerish;
        reasons.push(`Header-like first row (+${scoreWeights.headerish})`);
      }

      const mergedArea = sheet.merges.filter((merge) => {
        return !(merge.endRow < startRow || merge.startRow > endRow || merge.endCol < startCol || merge.startCol > endCol);
      }).length;
      if (mergedArea >= Math.max(2, Math.ceil(area * 0.08))) {
        score += scoreWeights.mergeHeavyPenalty;
        reasons.push(`Many merged cells (${scoreWeights.mergeHeavyPenalty})`);
      }

      if (sourceKind === "border") {
        if (mergedArea >= 2 && density < 0.25 && headerishCount < 2) {
          return;
        }
      } else if (mergedArea >= 2 && rowCount <= 6 && colCount >= 10 && density < 0.25) {
        return;
      }

      const nonEmptyCells = component.filter((entry) => entry.outputValue.trim());
      const avgTextLength = nonEmptyCells
        .reduce((sum, entry) => sum + entry.outputValue.trim().length, 0) / Math.max(1, nonEmptyCells.length);
      if (avgTextLength > 36 && density < 0.7) {
        score += scoreWeights.prosePenalty;
        reasons.push(`Mostly long prose (${scoreWeights.prosePenalty})`);
      }

      if (score >= scoreWeights.threshold) {
        const normalizedBounds = trimTableCandidateBounds(cellMap, {
          startRow,
          startCol,
          endRow,
          endCol
        });
        const key = `${normalizedBounds.startRow}:${normalizedBounds.startCol}:${normalizedBounds.endRow}:${normalizedBounds.endCol}`;
        if (candidateKeys.has(key)) {
          return;
        }
        candidateKeys.add(key);
        candidates.push({
          startRow: normalizedBounds.startRow,
          startCol: normalizedBounds.startCol,
          endRow: normalizedBounds.endRow,
          endCol: normalizedBounds.endCol,
          score,
          reasonSummary: reasons
        });
      }
    }

    for (const component of collectConnectedComponents(borderSeedCells, tableDetectionMode === "border" ? "border" : "grid")) {
      maybePushCandidate(component, "border");
    }

    if (tableDetectionMode !== "border") {
      for (const component of collectConnectedComponents(allSeedCells)) {
        const rows = component.map((entry) => entry.row);
        const cols = component.map((entry) => entry.col);
        const bounds = {
          startRow: Math.min(...rows),
          startCol: Math.min(...cols),
          endRow: Math.max(...rows),
          endCol: Math.max(...cols)
        };
        const containingBorderCandidates = candidates.filter((candidate) => isWithinBounds(candidate, bounds));
        const fallbackArea = getBoundsArea(bounds);
        const shadowedByBorderCandidate = containingBorderCandidates.some((candidate) => (
          getBoundsArea(candidate) >= fallbackArea * 0.4
        ));
        const shadowedByMultipleBorderCandidates = containingBorderCandidates.length >= 2
          && getCombinedCandidateArea(containingBorderCandidates) >= fallbackArea * 0.6;
        if (shadowedByBorderCandidate || shadowedByMultipleBorderCandidates) {
          continue;
        }
        maybePushCandidate(component, "fallback");
      }
    }

    return pruneRedundantCandidates(candidates).sort((left, right) => {
      if (left.startRow !== right.startRow) return left.startRow - right.startRow;
      return left.startCol - right.startCol;
    });
  }

  function trimTableCandidateBounds<TCell extends CellLike>(
    cellMap: Map<string, TCell>,
    bounds: { startRow: number; startCol: number; endRow: number; endCol: number }
  ): { startRow: number; startCol: number; endRow: number; endCol: number } {
    let { startRow, startCol, endRow, endCol } = bounds;
    const minBorderedCells = Math.max(2, Math.ceil((endCol - startCol + 1) * 0.5));

    while (endRow - startRow + 1 >= 2) {
      const topStats = borderGridHelper.collectTableEdgeStats(cellMap, startRow, startCol, endCol);
      const nextStats = borderGridHelper.collectTableEdgeStats(cellMap, startRow + 1, startCol, endCol);
      const shouldTrimTop = (
        topStats.nonEmptyCount <= 2
        && topStats.rawBorderCount === 0
        && nextStats.borderCount >= minBorderedCells
        && nextStats.nonEmptyCount >= Math.max(2, Math.ceil((endCol - startCol + 1) * 0.5))
      );
      if (!shouldTrimTop) {
        break;
      }
      startRow += 1;
    }

    for (let row = startRow + 1; row <= endRow; row += 1) {
      const currentStats = borderGridHelper.collectTableEdgeStats(cellMap, row, startCol, endCol);
      const previousStats = borderGridHelper.collectTableEdgeStats(cellMap, row - 1, startCol, endCol);
      const shouldBreakAtCurrentRow = (
        (previousStats.borderCount >= minBorderedCells
          || previousStats.bottomCount >= minBorderedCells
          || currentStats.topCount >= minBorderedCells)
        && currentStats.rawBorderCount === 0
        && currentStats.nonEmptyCount <= 1
      );
      if (shouldBreakAtCurrentRow) {
        endRow = row - 1;
        break;
      }
    }

    while (endRow - startRow + 1 >= 2) {
      const bottomStats = borderGridHelper.collectTableEdgeStats(cellMap, endRow, startCol, endCol);
      const previousStats = borderGridHelper.collectTableEdgeStats(cellMap, endRow - 1, startCol, endCol);
      const shouldTrimBottom = (
        (previousStats.borderCount >= minBorderedCells
          || previousStats.bottomCount >= minBorderedCells
          || bottomStats.topCount >= minBorderedCells)
        && bottomStats.rawBorderCount === 0
        && bottomStats.nonEmptyCount <= 1
      ) || (
        bottomStats.nonEmptyCount <= 1
        && bottomStats.rawBorderCount === 0
        && bottomStats.maxTextLength >= 12
        && previousStats.nonEmptyCount >= Math.max(2, Math.ceil((endCol - startCol + 1) * 0.5))
      );
      if (!shouldTrimBottom) {
        break;
      }
      endRow -= 1;
    }

    return { startRow, startCol, endRow, endCol };
  }

  function matrixFromCandidate<TCell extends CellLike>(
    sheet: SheetLike<TCell>,
    candidate: TableCandidate,
    options: MarkdownOptions,
    buildCellMap: (sheet: SheetLike<TCell>) => Map<string, TCell>,
    formatCellForMarkdown: (cell: TCell | undefined, options: MarkdownOptions) => string
  ): string[][] {
    const cellMap = buildCellMap(sheet);
    const rows: string[][] = [];
    for (let row = candidate.startRow; row <= candidate.endRow; row += 1) {
      const currentRow: string[] = [];
      for (let col = candidate.startCol; col <= candidate.endCol; col += 1) {
        const cell = cellMap.get(`${row}:${col}`);
        let value = formatCellForMarkdown(cell, options);
        if (options.trimText !== false) {
          value = value.trim();
        }
        currentRow.push(value);
      }
      rows.push(currentRow);
    }
    applyMergeTokens(rows, sheet.merges, candidate.startRow, candidate.startCol, candidate.endRow, candidate.endCol);

    let normalizedRows = rows;
    if (options.removeEmptyRows !== false) {
      normalizedRows = normalizedRows.filter((row) => row.some((cell) => isMeaningfulMarkdownCell(cell)));
    }
    if (options.removeEmptyColumns !== false && normalizedRows.length > 0) {
      const keepColumnFlags = normalizedRows[0].map((_, colIndex) => normalizedRows.some((row) => isMeaningfulMarkdownCell(row[colIndex])));
      normalizedRows = normalizedRows.map((row) => row.filter((_cell, colIndex) => keepColumnFlags[colIndex]));
    }
    return normalizedRows;
  }

  function isMeaningfulMarkdownCell(value: string): boolean {
    const text = String(value || "").trim();
    if (!text) return false;
    return text !== "[MERGED←]" && text !== "[MERGED↑]";
  }

  function applyMergeTokens(
    matrix: string[][],
    merges: MergeRange[],
    startRow: number,
    startCol: number,
    endRow: number,
    endCol: number
  ): void {
    for (const merge of merges) {
      if (merge.endRow < startRow || merge.startRow > endRow || merge.endCol < startCol || merge.startCol > endCol) {
        continue;
      }
      for (let row = merge.startRow; row <= merge.endRow; row += 1) {
        for (let col = merge.startCol; col <= merge.endCol; col += 1) {
          if (row === merge.startRow && col === merge.startCol) continue;
          const matrixRow = row - startRow;
          const matrixCol = col - startCol;
          if (!matrix[matrixRow] || typeof matrix[matrixRow][matrixCol] === "undefined") {
            continue;
          }
          matrix[matrixRow][matrixCol] = row === merge.startRow ? "[MERGED←]" : "[MERGED↑]";
        }
      }
    }
  }

  const tableDetectorApi = {
    collectTableSeedCells,
    collectBorderSeedCells,
    pruneRedundantCandidates,
    detectTableCandidates,
    trimTableCandidateBounds,
    matrixFromCandidate,
    isMeaningfulMarkdownCell,
    applyMergeTokens,
    defaultTableScoreWeights: DEFAULT_TABLE_SCORE_WEIGHTS
  };

  moduleRegistry.registerModule("tableDetector", tableDetectorApi);
})();
