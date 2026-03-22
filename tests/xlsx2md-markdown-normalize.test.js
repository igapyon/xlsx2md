// @vitest-environment jsdom

import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";
import { loadModuleRegistry } from "./helpers/module-registry.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const markdownNormalizeCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/markdown-normalize.js"),
  "utf8"
);

function bootMarkdownNormalize() {
  document.body.innerHTML = "";
  loadModuleRegistry(__dirname);
  new Function(markdownNormalizeCode)();
  return globalThis.__xlsx2mdModuleRegistry.getModule("markdownNormalize");
}

describe("xlsx2md markdown normalize", () => {
  it("replaces line breaks, tabs, and control characters with spaces without trimming", () => {
    const api = bootMarkdownNormalize();

    expect(api.normalizeMarkdownText(" a\r\nb\tc\u0007d ")).toBe(" a b c d ");
  });

  it("replaces unsafe unicode format and separator characters with spaces", () => {
    const api = bootMarkdownNormalize();

    expect(api.normalizeMarkdownText("a\u0085b\u200Bc\u2028d\u202Ee\u2060f\uFEFFg\u00ADh")).toBe("a b c d e f g h");
  });

  it("escapes pipes in table cells after normalization", () => {
    const api = bootMarkdownNormalize();

    expect(api.normalizeMarkdownTableCell("A|\nB")).toBe("A\\| B");
  });

  it("preserves surrounding spaces while normalizing tabs and pipes in table cells", () => {
    const api = bootMarkdownNormalize();

    expect(api.normalizeMarkdownTableCell("  A\t|\tB  ")).toBe("  A \\| B  ");
  });

  it("removes heading and list markers from normalized text", () => {
    const api = bootMarkdownNormalize();

    expect(api.normalizeMarkdownHeadingText("### heading")).toBe("heading");
    expect(api.normalizeMarkdownListItemText("- item")).toBe("item");
    expect(api.normalizeMarkdownListItemText("1. item")).toBe("item");
  });
});
