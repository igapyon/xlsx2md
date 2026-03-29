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
  path.resolve(__dirname, "../src/js/markdown-escape.js"),
  "utf8"
);
const richTextParserCode = readFileSync(
  path.resolve(__dirname, "../src/js/rich-text-parser.js"),
  "utf8"
);

function bootRichTextParser() {
  document.body.innerHTML = "";
  loadModuleRegistry(__dirname);
  new Function(markdownEscapeCode)();
  new Function(richTextParserCode)();
  return globalThis.__xlsx2mdModuleRegistry.getModule("richTextParser")
    .createRichTextParserApi({
      normalizeMarkdownText: (text) => String(text || "").replace(/\r\n?|\n/g, " ").replace(/\t/g, " ")
    });
}

describe("xlsx2md rich text parser", () => {
  it("tokenizes plain mode into a single escaped text token", () => {
    const api = bootRichTextParser();

    expect(api.tokenizeCellDisplayText({
      outputValue: "# a*b\nc",
      textStyle: { bold: false, italic: false, strike: false, underline: false },
      richTextRuns: null
    }, "plain")).toEqual([
      { kind: "text", text: "\\# a\\*b c" }
    ]);
  });

  it("tokenizes github fallback cells into styledText and lineBreak tokens", () => {
    const api = bootRichTextParser();

    expect(api.tokenizeCellDisplayText({
      outputValue: "line1\nline2",
      textStyle: { bold: true, italic: false, strike: false, underline: true },
      richTextRuns: null
    }, "github")).toEqual([
      {
        kind: "styledText",
        parts: [{ kind: "text", text: "line1", rawText: "line1" }],
        style: { bold: true, italic: false, strike: false, underline: true }
      },
      { kind: "lineBreak" },
      {
        kind: "styledText",
        parts: [{ kind: "text", text: "line2", rawText: "line2" }],
        style: { bold: true, italic: false, strike: false, underline: true }
      }
    ]);
  });

  it("tokenizes github rich runs with escaped markdown text and preserved run boundaries", () => {
    const api = bootRichTextParser();

    expect(api.tokenizeCellDisplayText({
      outputValue: "a*b plain",
      textStyle: { bold: false, italic: false, strike: false, underline: false },
      richTextRuns: [
        { text: "a*b ", bold: true, italic: false, strike: false, underline: false },
        { text: "plain", bold: false, italic: false, strike: false, underline: false }
      ]
    }, "github")).toEqual([
      {
        kind: "styledText",
        parts: [
          { kind: "text", text: "a", rawText: "a" },
          { kind: "escaped", text: "\\*", rawText: "*" },
          { kind: "text", text: "b ", rawText: "b " }
        ],
        style: { bold: true, italic: false, strike: false, underline: false }
      },
      {
        kind: "styledText",
        parts: [{ kind: "text", text: "plain", rawText: "plain" }],
        style: { bold: false, italic: false, strike: false, underline: false }
      }
    ]);
  });

  it("preserves run-boundary spaces while escaping markdown literals", () => {
    const api = bootRichTextParser();

    expect(api.tokenizeCellDisplayText({
      outputValue: "plain bold text",
      textStyle: { bold: false, italic: false, strike: false, underline: false },
      richTextRuns: [
        { text: "plain ", bold: false, italic: false, strike: false, underline: false },
        { text: "bold", bold: true, italic: false, strike: false, underline: false },
        { text: " text", bold: false, italic: false, strike: false, underline: false }
      ]
    }, "github")).toEqual([
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
      {
        kind: "styledText",
        parts: [{ kind: "text", text: " text", rawText: " text" }],
        style: { bold: false, italic: false, strike: false, underline: false }
      }
    ]);
  });

  it("tokenizes image-like markdown text into escaped parts inside styled runs", () => {
    const api = bootRichTextParser();

    expect(api.tokenizeCellDisplayText({
      outputValue: "![alt](img.png)",
      textStyle: { bold: false, italic: false, strike: false, underline: false },
      richTextRuns: [
        { text: "![alt](img.png)", bold: false, italic: false, strike: false, underline: true }
      ]
    }, "github")).toEqual([
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
        style: { bold: false, italic: false, strike: false, underline: true }
      }
    ]);
  });

  it("shows plain-vs-github token differences for the same fallback cell", () => {
    const api = bootRichTextParser();
    const cell = {
      outputValue: "a*b\nnext",
      textStyle: { bold: true, italic: false, strike: false, underline: false },
      richTextRuns: null
    };

    expectModeResults(
      (mode) => api.tokenizeCellDisplayText(cell, mode),
      {
        plain: [
          { kind: "text", text: "a\\*b next" }
        ],
        github: [
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
        ]
      }
    );
  });

  it("shows plain-vs-github token differences for the same escaped line-start text", () => {
    const api = bootRichTextParser();
    const cell = {
      outputValue: "# head",
      textStyle: { bold: false, italic: true, strike: false, underline: false },
      richTextRuns: null
    };

    expectModeResults(
      (mode) => api.tokenizeCellDisplayText(cell, mode),
      {
        plain: [
          { kind: "text", text: "\\# head" }
        ],
        github: [
          {
            kind: "styledText",
            parts: [
              { kind: "escaped", text: "\\#", rawText: "#" },
              { kind: "text", text: " head", rawText: " head" }
            ],
            style: { bold: false, italic: true, strike: false, underline: false }
          }
        ]
      }
    );
  });
});
