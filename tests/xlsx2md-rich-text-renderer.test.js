// @vitest-environment jsdom

import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";
import { expectModeResults } from "./helpers/mode-assertions.js";
import { loadModuleRegistry } from "./helpers/module-registry.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const markdownEscapeCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/markdown-escape.js"),
  "utf8"
);
const richTextParserCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/rich-text-parser.js"),
  "utf8"
);
const richTextPlainFormatterCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/rich-text-plain-formatter.js"),
  "utf8"
);
const richTextGithubFormatterCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/rich-text-github-formatter.js"),
  "utf8"
);
const richTextRendererCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/rich-text-renderer.js"),
  "utf8"
);

function bootRichTextRenderer() {
  document.body.innerHTML = "";
  loadModuleRegistry(__dirname);
  new Function(markdownEscapeCode)();
  new Function(richTextParserCode)();
  new Function(richTextPlainFormatterCode)();
  new Function(richTextGithubFormatterCode)();
  new Function(richTextRendererCode)();
  return globalThis.__xlsx2mdModuleRegistry.getModule("richTextRenderer")
    .createRichTextRendererApi({
      normalizeMarkdownText: (text) => String(text || "").replace(/\r\n?|\n/g, " ").replace(/\t/g, " ")
    });
}

describe("xlsx2md rich text renderer", () => {
  it("tokenizes rich runs into styledText and lineBreak tokens for github mode", () => {
    const api = bootRichTextRenderer();

    expect(api.tokenizeGithubRichTextRuns([
      { text: "plain ", bold: false, italic: false, strike: false, underline: false },
      { text: "bold\nnext", bold: true, italic: false, strike: false, underline: false }
    ])).toEqual([
      {
        kind: "styledText",
        parts: [{ kind: "text", text: "plain ", rawText: "plain " }],
        style: { bold: false, italic: false, strike: false, underline: false }
      },
      {
        kind: "styledText",
        parts: [{ kind: "text", text: "bold", rawText: "bold" }],
        style: { bold: true, italic: false, strike: false, underline: false }
      },
      { kind: "lineBreak" },
      {
        kind: "styledText",
        parts: [{ kind: "text", text: "next", rawText: "next" }],
        style: { bold: true, italic: false, strike: false, underline: false }
      }
    ]);
  });

  it("renders tokenized rich text with github formatting", () => {
    const api = bootRichTextRenderer();
    const cell = {
      outputValue: "plain bold\nnext",
      textStyle: { bold: false, italic: false, strike: false, underline: false },
      richTextRuns: [
        { text: "plain ", bold: false, italic: false, strike: false, underline: false },
        { text: "bold\nnext", bold: true, italic: false, strike: false, underline: false }
      ]
    };

    expect(api.renderCellDisplayText(cell, "github")).toBe("plain **bold**<br>**next**");
  });

  it("falls back to plain tokens in plain mode", () => {
    const api = bootRichTextRenderer();

    expect(api.tokenizeCellDisplayText({
      outputValue: "# head\nline2",
      textStyle: { bold: true, italic: false, strike: false, underline: false },
      richTextRuns: null
    }, "plain")).toEqual([
      { kind: "text", text: "\\# head line2" }
    ]);
    expect(api.renderCellDisplayText({
      outputValue: "# head\nline2",
      textStyle: { bold: true, italic: false, strike: false, underline: false },
      richTextRuns: null
    }, "plain")).toBe("\\# head line2");
  });

  it("renders styled text parts before applying wrappers", () => {
    const api = bootRichTextRenderer();

    expect(api.renderStyledTextParts([
      { kind: "text", text: "a\\*b" },
      { kind: "text", text: " c" }
    ])).toBe("a\\*b c");
  });

  it("delegates github-token rendering to the github formatter layer", () => {
    const api = bootRichTextRenderer();

    expect(api.renderGithubTokens([
      { kind: "text", text: "plain" },
      { kind: "lineBreak" },
      {
        kind: "styledText",
        parts: [{ kind: "text", text: "x" }],
        style: { bold: false, italic: true, strike: false, underline: false }
      }
    ])).toBe("plain<br>*x*");
  });

  it("delegates plain-token rendering to the plain formatter layer", () => {
    const api = bootRichTextRenderer();

    expect(api.renderPlainTokens([
      { kind: "text", text: "\\# h" },
      { kind: "lineBreak" },
      {
        kind: "styledText",
        parts: [{ kind: "text", text: "x" }],
        style: { bold: true, italic: false, strike: false, underline: false }
      }
    ])).toBe("\\# h x");
  });

  it("keeps escaped list markers literal inside github-styled text", () => {
    const api = bootRichTextRenderer();
    const cell = {
      outputValue: "- item",
      textStyle: { bold: true, italic: false, strike: false, underline: false },
      richTextRuns: null
    };

    expect(api.renderCellDisplayText(cell, "github")).toBe("**\\- item**");
  });

  it("keeps escaped ordered-list and quote markers literal inside github-styled text", () => {
    const api = bootRichTextRenderer();

    expect(api.renderCellDisplayText({
      outputValue: "1. item",
      textStyle: { bold: false, italic: true, strike: false, underline: false },
      richTextRuns: null
    }, "github")).toBe("*1\\. item*");

    expect(api.renderCellDisplayText({
      outputValue: "> quote",
      textStyle: { bold: false, italic: false, strike: true, underline: false },
      richTextRuns: null
    }, "github")).toBe("~~&gt; quote~~");
  });

  it("renders escaped markdown link-like text safely inside styled runs", () => {
    const api = bootRichTextRenderer();
    const cell = {
      outputValue: "[x](y) `code` <tag>",
      textStyle: { bold: false, italic: false, strike: false, underline: false },
      richTextRuns: [
        { text: "[x](y) ", bold: true, italic: false, strike: false, underline: false },
        { text: "`code` ", bold: false, italic: true, strike: false, underline: false },
        { text: "<tag>", bold: false, italic: false, strike: false, underline: true }
      ]
    };

    expect(api.renderCellDisplayText(cell, "github")).toBe("**\\[x\\]\\(y\\) ***\\`code\\` *<ins>&lt;tag&gt;</ins>");
  });

  it("renders escaped ampersands and image-like text across styled rich runs", () => {
    const api = bootRichTextRenderer();
    const cell = {
      outputValue: "a & b ![alt](img.png)",
      textStyle: { bold: false, italic: false, strike: false, underline: false },
      richTextRuns: [
        { text: "a & b ", bold: true, italic: false, strike: false, underline: false },
        { text: "![alt](img.png)", bold: false, italic: true, strike: true, underline: true }
      ]
    };

    expect(api.renderCellDisplayText(cell, "github")).toBe("**a &amp; b ***~~<ins>\\!\\[alt\\]\\(img.png\\)</ins>~~*");
  });

  it("renders consecutive line breaks across styled rich runs", () => {
    const api = bootRichTextRenderer();
    const cell = {
      outputValue: "top\n\nnext",
      textStyle: { bold: false, italic: false, strike: false, underline: false },
      richTextRuns: [
        { text: "top\n", bold: true, italic: false, strike: false, underline: false },
        { text: "\nnext", bold: false, italic: true, strike: false, underline: true }
      ]
    };

    expect(api.renderCellDisplayText(cell, "github")).toBe("**top**<br><br>*<ins>next</ins>*");
  });

  it("renders plus and star markers literally across styled rich runs", () => {
    const api = bootRichTextRenderer();
    const cell = {
      outputValue: "+ plus * star",
      textStyle: { bold: false, italic: false, strike: false, underline: false },
      richTextRuns: [
        { text: "+ plus ", bold: true, italic: false, strike: false, underline: false },
        { text: "* star", bold: false, italic: true, strike: false, underline: false }
      ]
    };

    expect(api.renderCellDisplayText(cell, "github")).toBe("**\\+ plus ***\\* star*");
  });

  it("shows plain-vs-github differences for the same rich-text input", () => {
    const api = bootRichTextRenderer();
    const cell = {
      outputValue: "a*b\nnext",
      textStyle: { bold: true, italic: false, strike: false, underline: false },
      richTextRuns: null
    };

    expectModeResults(
      (mode) => api.renderCellDisplayText(cell, mode),
      {
        plain: "a\\*b next",
        github: "**a\\*b**<br>**next**"
      }
    );
  });
});
