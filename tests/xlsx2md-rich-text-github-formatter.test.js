// @vitest-environment jsdom

import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";
import { loadModuleRegistry } from "./helpers/module-registry.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const richTextGithubFormatterCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/rich-text-github-formatter.js"),
  "utf8"
);

function bootRichTextGithubFormatter() {
  document.body.innerHTML = "";
  loadModuleRegistry(__dirname);
  new Function(richTextGithubFormatterCode)();
  return globalThis.__xlsx2mdModuleRegistry.getModule("richTextGithubFormatter")
    .createRichTextGithubFormatterApi();
}

describe("xlsx2md rich text github formatter", () => {
  it("applies github-compatible wrappers in stable order", () => {
    const api = bootRichTextGithubFormatter();

    expect(api.applyTextStyle("text", {
      bold: true,
      italic: true,
      strike: true,
      underline: true
    })).toBe("***~~<ins>text</ins>~~***");
  });

  it("renders github tokens with <br> and styled parts", () => {
    const api = bootRichTextGithubFormatter();

    expect(api.renderGithubTokens([
      { kind: "text", text: "plain " },
      {
        kind: "styledText",
        parts: [{ kind: "text", text: "a\\*b", rawText: "a*b" }],
        style: { bold: true, italic: false, strike: false, underline: false }
      },
      { kind: "lineBreak" },
      {
        kind: "styledText",
        parts: [{ kind: "text", text: "next", rawText: "next" }],
        style: { bold: false, italic: true, strike: false, underline: false }
      }
    ])).toBe("plain **a\\*b**<br>*next*");
  });

  it("renders escaped parts via the escaped-part branch", () => {
    const api = bootRichTextGithubFormatter();

    expect(api.renderStyledTextPart({ kind: "escaped", text: "\\*", rawText: "*" })).toBe("\\*");
    expect(api.renderStyledTextPart({ kind: "text", text: "a", rawText: "a" })).toBe("a");
  });
});
