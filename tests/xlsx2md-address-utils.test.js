// @vitest-environment jsdom

import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";
import { loadModuleRegistry } from "./helpers/module-registry.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const addressUtilsCode = readFileSync(
  path.resolve(__dirname, "../src/js/address-utils.js"),
  "utf8"
);

function bootAddressUtils() {
  document.body.innerHTML = "";
  loadModuleRegistry(__dirname);
  new Function(addressUtilsCode)();
  return globalThis.__xlsx2mdModuleRegistry.getModule("addressUtils");
}

describe("xlsx2md address utils", () => {
  it("converts between column numbers and letters", () => {
    const api = bootAddressUtils();

    expect(api.colToLetters(1)).toBe("A");
    expect(api.colToLetters(28)).toBe("AB");
    expect(api.lettersToCol("AB")).toBe(28);
  });

  it("parses and normalizes cell and range addresses", () => {
    const api = bootAddressUtils();

    expect(api.parseCellAddress("$C$12")).toEqual({ row: 12, col: 3 });
    expect(api.normalizeFormulaAddress("$d$7")).toBe("D7");
    expect(api.parseRangeAddress("$A$1:$C$3")).toEqual({ start: "A1", end: "C3" });
  });

  it("formats display ranges and merged range refs", () => {
    const api = bootAddressUtils();

    expect(api.formatRange(12, 2, 17, 6)).toBe("B12-F17");
    expect(api.parseRangeRef("B12:F17")).toEqual({
      startRow: 12,
      startCol: 2,
      endRow: 17,
      endCol: 6,
      ref: "B12:F17"
    });
  });
});
