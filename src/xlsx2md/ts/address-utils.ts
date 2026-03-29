/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */

(() => {
  const moduleRegistry = getXlsx2mdModuleRegistry();
  type MergeRange = {
    startRow: number;
    startCol: number;
    endRow: number;
    endCol: number;
    ref: string;
  };

  function colToLetters(col: number): string {
    let current = col;
    let result = "";
    while (current > 0) {
      const remainder = (current - 1) % 26;
      result = String.fromCharCode(65 + remainder) + result;
      current = Math.floor((current - 1) / 26);
    }
    return result;
  }

  function lettersToCol(letters: string): number {
    let result = 0;
    for (const ch of String(letters || "").toUpperCase()) {
      result = result * 26 + (ch.charCodeAt(0) - 64);
    }
    return result;
  }

  function parseCellAddress(address: string): { row: number; col: number } {
    const normalized = String(address || "").trim().replace(/\$/g, "");
    const match = normalized.match(/^([A-Z]+)(\d+)$/i);
    if (!match) {
      return { row: 0, col: 0 };
    }
    return {
      col: lettersToCol(match[1]),
      row: Number(match[2])
    };
  }

  function normalizeFormulaAddress(address: string): string {
    return String(address || "").trim().replace(/\$/g, "").toUpperCase();
  }

  function formatRange(startRow: number, startCol: number, endRow: number, endCol: number): string {
    return `${colToLetters(startCol)}${startRow}-${colToLetters(endCol)}${endRow}`;
  }

  function parseRangeRef(ref: string): MergeRange {
    const parts = String(ref || "").split(":");
    const start = parseCellAddress(parts[0] || "");
    const end = parseCellAddress(parts[1] || parts[0] || "");
    return {
      startRow: start.row,
      startCol: start.col,
      endRow: end.row,
      endCol: end.col,
      ref
    };
  }

  function parseRangeAddress(rawRange: string): { start: string; end: string } | null {
    const match = String(rawRange || "").trim().match(/^(\$?[A-Z]+\$?\d+):(\$?[A-Z]+\$?\d+)$/i);
    if (!match) return null;
    return {
      start: normalizeFormulaAddress(match[1]),
      end: normalizeFormulaAddress(match[2])
    };
  }

  const addressUtilsApi = {
    colToLetters,
    lettersToCol,
    parseCellAddress,
    normalizeFormulaAddress,
    formatRange,
    parseRangeRef,
    parseRangeAddress
  };

  moduleRegistry.registerModule("addressUtils", addressUtilsApi);
})();
