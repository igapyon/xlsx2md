// @vitest-environment jsdom

import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";
import { loadModuleRegistry } from "./helpers/module-registry.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const richTextGithubFormatterCode = readFileSync(
  path.resolve(__dirname, "../src/js/rich-text-github-formatter.js"),
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

  it("renders line breaks across combined styles in stable wrapper order", () => {
    const api = bootRichTextGithubFormatter();

    expect(api.renderGithubTokens([
      {
        kind: "styledText",
        parts: [{ kind: "text", text: "top", rawText: "top" }],
        style: { bold: true, italic: false, strike: false, underline: true }
      },
      { kind: "lineBreak" },
      {
        kind: "styledText",
        parts: [{ kind: "text", text: "next", rawText: "next" }],
        style: { bold: false, italic: true, strike: true, underline: false }
      }
    ])).toBe("**<ins>top</ins>**<br>*~~next~~*");
  });

  it("collapses repeated spaces after joining github tokens", () => {
    const api = bootRichTextGithubFormatter();

    expect(api.renderGithubTokens([
      { kind: "text", text: "a  " },
      {
        kind: "styledText",
        parts: [{ kind: "text", text: "b", rawText: "b" }],
        style: { bold: true, italic: false, strike: false, underline: false }
      },
      { kind: "text", text: "   c" }
    ])).toBe("a **b** c");
  });

  it("renders escaped ampersands and image-like text inside combined styles", () => {
    const api = bootRichTextGithubFormatter();

    expect(api.renderGithubTokens([
      {
        kind: "styledText",
        parts: [
          { kind: "text", text: "a ", rawText: "a " },
          { kind: "escaped", text: "&amp;", rawText: "&" },
          { kind: "text", text: " b", rawText: " b" }
        ],
        style: { bold: true, italic: false, strike: false, underline: false }
      },
      { kind: "text", text: " " },
      {
        kind: "styledText",
        parts: [
          { kind: "escaped", text: "\\!", rawText: "!" },
          { kind: "escaped", text: "\\[", rawText: "[" },
          { kind: "text", text: "alt", rawText: "alt" },
          { kind: "escaped", text: "\\]", rawText: "]" },
          { kind: "escaped", text: "\\(", rawText: "(" },
          { kind: "text", text: "img.png", rawText: "img.png" },
          { kind: "escaped", text: "\\)", rawText: ")" }
        ],
        style: { bold: false, italic: true, strike: true, underline: true }
      }
    ])).toBe("**a &amp; b** *~~<ins>\\!\\[alt\\]\\(img.png\\)</ins>~~*");
  });

  it("renders escaped plus and star markers literally inside styled runs", () => {
    const api = bootRichTextGithubFormatter();

    expect(api.renderGithubTokens([
      {
        kind: "styledText",
        parts: [
          { kind: "escaped", text: "\\+", rawText: "+" },
          { kind: "text", text: " plus", rawText: " plus" }
        ],
        style: { bold: true, italic: false, strike: false, underline: false }
      },
      { kind: "text", text: " " },
      {
        kind: "styledText",
        parts: [
          { kind: "escaped", text: "\\*", rawText: "*" },
          { kind: "text", text: " star", rawText: " star" }
        ],
        style: { bold: false, italic: true, strike: false, underline: false }
      }
    ])).toBe("**\\+ plus** *\\* star*");
  });

  it("shows formatter-level difference from plain rendering for the same tokens", () => {
    const api = bootRichTextGithubFormatter();
    const tokens = [
      {
        kind: "styledText",
        parts: [
          { kind: "text", text: "a", rawText: "a" },
          { kind: "escaped", text: "\\*", rawText: "*" },
          { kind: "text", text: "b", rawText: "b" }
        ],
        style: { bold: true, italic: false, strike: false, underline: false }
      },
      { kind: "lineBreak" },
      {
        kind: "styledText",
        parts: [{ kind: "text", text: "next", rawText: "next" }],
        style: { bold: true, italic: false, strike: false, underline: false }
      }
    ];

    expect(api.renderGithubTokens(tokens)).toBe("**a\\*b**<br>**next**");
  });

  it("renders escaped line-start markers inside styled github tokens", () => {
    const api = bootRichTextGithubFormatter();

    expect(api.renderGithubTokens([
      {
        kind: "styledText",
        parts: [
          { kind: "escaped", text: "\\#", rawText: "#" },
          { kind: "text", text: " head", rawText: " head" }
        ],
        style: { bold: false, italic: true, strike: false, underline: false }
      }
    ])).toBe("*\\# head*");
  });
});
