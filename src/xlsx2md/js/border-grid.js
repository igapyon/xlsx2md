(() => {
    const moduleRegistry = getXlsx2mdModuleRegistry();
    function getCellAt(cellMap, row, col) {
        return cellMap.get(`${row}:${col}`);
    }
    function hasNormalizedBorderOnSide(cellMap, row, col, side) {
        const cell = getCellAt(cellMap, row, col);
        if (side === "top") {
            const above = getCellAt(cellMap, row - 1, col);
            return !!(cell === null || cell === void 0 ? void 0 : cell.borders.top) || !!(above === null || above === void 0 ? void 0 : above.borders.bottom);
        }
        if (side === "bottom") {
            const below = getCellAt(cellMap, row + 1, col);
            return !!(cell === null || cell === void 0 ? void 0 : cell.borders.bottom) || !!(below === null || below === void 0 ? void 0 : below.borders.top);
        }
        if (side === "left") {
            const left = getCellAt(cellMap, row, col - 1);
            return !!(cell === null || cell === void 0 ? void 0 : cell.borders.left) || !!(left === null || left === void 0 ? void 0 : left.borders.right);
        }
        const right = getCellAt(cellMap, row, col + 1);
        return !!(cell === null || cell === void 0 ? void 0 : cell.borders.right) || !!(right === null || right === void 0 ? void 0 : right.borders.left);
    }
    function hasAnyNormalizedBorder(cellMap, row, col) {
        return hasNormalizedBorderOnSide(cellMap, row, col, "top")
            || hasNormalizedBorderOnSide(cellMap, row, col, "bottom")
            || hasNormalizedBorderOnSide(cellMap, row, col, "left")
            || hasNormalizedBorderOnSide(cellMap, row, col, "right");
    }
    function collectTableEdgeStats(cellMap, row, startCol, endCol) {
        let nonEmptyCount = 0;
        let borderCount = 0;
        let rawBorderCount = 0;
        let topCount = 0;
        let bottomCount = 0;
        let maxTextLength = 0;
        for (let col = startCol; col <= endCol; col += 1) {
            const cell = getCellAt(cellMap, row, col);
            const text = String((cell === null || cell === void 0 ? void 0 : cell.outputValue) || "").trim();
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
    function countNormalizedBorderedCells(cellMap, startRow, startCol, endRow, endCol) {
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
