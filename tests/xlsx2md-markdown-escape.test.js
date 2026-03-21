// @vitest-environment jsdom

import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";
import { loadModuleRegistry } from "./helpers/module-registry.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const markdownEscapeCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/markdown-escape.js"),
  "utf8"
);

function bootMarkdownEscape() {
  document.body.innerHTML = "";
  loadModuleRegistry(__dirname);
  new Function(markdownEscapeCode)();
  return globalThis.__xlsx2mdModuleRegistry.getModule("markdownEscape");
}

describe("xlsx2md markdown escape", () => {
  it("escapes inline markdown control characters and html-sensitive brackets", () => {
    const api = bootMarkdownEscape();

    expect(api.escapeMarkdownLiteralText("a*b _c_ [x](y) ![z](w) <tag> a | b `q` ~"))
      .toBe("a\\*b \\_c\\_ \\[x\\]\\(y\\) \\!\\[z\\]\\(w\\) &lt;tag&gt; a \\| b \\`q\\` \\~");
    expect(api.escapeMarkdownLiteralParts("a*b <tag>")).toEqual([
      { kind: "text", text: "a", rawText: "a" },
      { kind: "escaped", text: "\\*", rawText: "*" },
      { kind: "text", text: "b ", rawText: "b " },
      { kind: "escaped", text: "&lt;", rawText: "<" },
      { kind: "text", text: "tag", rawText: "tag" },
      { kind: "escaped", text: "&gt;", rawText: ">" }
    ]);
  });

  it("escapes line-start markdown markers line by line", () => {
    const api = bootMarkdownEscape();

    expect(api.escapeMarkdownLiteralText("# h\n- item\n1. num\n> quote"))
      .toBe("\\# h\n\\- item\n1\\. num\n&gt; quote");
  });
});
