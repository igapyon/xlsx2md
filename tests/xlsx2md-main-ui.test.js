// @vitest-environment jsdom

import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it, vi } from "vitest";
import { loadModuleRegistry } from "./helpers/module-registry.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const mainCode = readFileSync(
  path.resolve(__dirname, "../src/js/main.js"),
  "utf8"
);
const textEncodingCode = readFileSync(
  path.resolve(__dirname, "../src/js/text-encoding.js"),
  "utf8"
);

function createDomFixture() {
  document.body.innerHTML = `
    <input id="xlsxFileInput" type="file" />
    <button id="convertBtn" type="button">Convert</button>
    <button id="downloadBtn" type="button">Download</button>
    <button id="exportZipBtn" type="button">ZIP</button>
    <input id="headerRowEnabled" type="checkbox" checked />
    <input id="trimTextEnabled" type="checkbox" checked />
    <input id="removeEmptyRowsEnabled" type="checkbox" checked />
    <input id="removeEmptyColumnsEnabled" type="checkbox" checked />
    <input id="includeShapeDetailsEnabled" type="checkbox" checked />
    <select id="outputModeSelect">
      <option value="display" selected>display</option>
      <option value="raw">raw</option>
      <option value="both">both</option>
    </select>
    <select id="encodingSelect">
      <option value="utf-8" selected>utf-8</option>
      <option value="shift_jis">shift_jis</option>
      <option value="utf-16le">utf-16le</option>
    </select>
    <select id="bomSelect">
      <option value="off" selected>off</option>
      <option value="on">on</option>
    </select>
    <select id="formattingModeSelect">
      <option value="plain">plain</option>
      <option value="github" selected>github</option>
    </select>
    <select id="tableDetectionModeSelect">
      <option value="balanced" selected>balanced</option>
      <option value="border">border</option>
    </select>
    <div id="outputModeNotice"></div>
    <div id="formattingModeNotice"></div>
    <div id="tableDetectionModeNotice"></div>
    <div id="encodingNotice"></div>
    <div id="bomNotice"></div>
    <div id="previewModeBanner" hidden></div>
    <div id="analysisSummary"></div>
    <div id="scoreSummary"></div>
    <div id="formulaSummary"></div>
    <div id="markdownPreview"></div>
    <pre id="markdownOutput"></pre>
    <div id="loadingOverlay"></div>
    <div id="errorAlert"></div>
    <div id="toast"></div>
  `;

  const outputModeSelect = document.getElementById("outputModeSelect");
  outputModeSelect.getValue = function getValue() {
    return this.value;
  };
  const formattingModeSelect = document.getElementById("formattingModeSelect");
  formattingModeSelect.getValue = function getValue() {
    return this.value;
  };
  const encodingSelect = document.getElementById("encodingSelect");
  encodingSelect.getValue = function getValue() {
    return this.value;
  };
  const bomSelect = document.getElementById("bomSelect");
  bomSelect.getValue = function getValue() {
    return this.value;
  };
  const tableDetectionModeSelect = document.getElementById("tableDetectionModeSelect");
  tableDetectionModeSelect.getValue = function getValue() {
    return this.value;
  };

  const markdownPreview = document.getElementById("markdownPreview");
  markdownPreview.setText = function setText(text) {
    this.dataset.rendered = text;
    this.textContent = text;
  };

  const loadingOverlay = document.getElementById("loadingOverlay");
  loadingOverlay.show = function show(text) {
    this.dataset.active = "true";
    if (text) this.dataset.text = text;
  };
  loadingOverlay.hide = function hide() {
    delete this.dataset.active;
  };

  const errorAlert = document.getElementById("errorAlert");
  errorAlert.show = function show(text) {
    this.dataset.message = text;
    this.textContent = text;
  };
  errorAlert.clear = function clear() {
    delete this.dataset.message;
    this.textContent = "";
  };

  const toast = document.getElementById("toast");
  toast.show = function show(text) {
    this.dataset.message = text;
  };
}

function createWorkbookFile() {
  return {
    fileName: "book_001_Sheet1.md",
    sheetName: "Sheet1",
    markdown: "# Sheet1",
    summary: {
      outputMode: "display",
      formattingMode: "github",
      tableDetectionMode: "balanced",
      tables: 1,
      narrativeBlocks: 1,
      merges: 0,
      images: 0,
      cells: 2,
      tableScores: [],
      formulaDiagnostics: []
    }
  };
}

function bootMain(overrides = {}, options = {}) {
  createDomFixture();
  const registry = loadModuleRegistry(__dirname);
  if (options.disableNodeRequire) {
    delete globalThis.__xlsx2mdNodeRequire;
  }
  new Function(textEncodingCode)();
  const api = {
    parseWorkbook: vi.fn(async () => ({ name: "book.xlsx", sheets: [{ name: "Sheet1", index: 1 }] })),
    convertWorkbookToMarkdownFiles: vi.fn(() => [createWorkbookFile()]),
    createSummaryText: vi.fn(() => "summary"),
    createCombinedMarkdownExportFile: vi.fn(() => ({ fileName: "book.md", content: "# combined" })),
    createCombinedMarkdownExportPayload: vi.fn(() => ({ fileName: "book.md", content: "# combined", data: new Uint8Array([35]), mimeType: "text/markdown;charset=utf-8" })),
    createWorkbookExportArchive: vi.fn(() => new Uint8Array([1, 2, 3])),
    ...overrides
  };
  registry.registerModule("xlsx2md", api);
  new Function(mainCode)();
  document.dispatchEvent(new Event("DOMContentLoaded"));
  return api;
}

async function flushAsyncWork() {
  await Promise.resolve();
  await new Promise((resolve) => window.setTimeout(resolve, 0));
}

describe("xlsx2md main ui", () => {
  it("initializes the screen with disabled download actions and display notice", () => {
    bootMain();

    expect(document.getElementById("downloadBtn").disabled).toBe(true);
    expect(document.getElementById("exportZipBtn").disabled).toBe(true);
    expect(document.getElementById("outputModeNotice").textContent).toContain("`display`");
    expect(document.getElementById("formattingModeNotice").textContent).toContain("`github`");
    expect(document.getElementById("tableDetectionModeNotice").textContent).toContain("`balanced`");
    expect(document.getElementById("encodingNotice").textContent).toContain("`utf-8`");
    expect(document.getElementById("bomNotice").textContent).toContain("disabled");
    expect(document.getElementById("analysisSummary").textContent).toContain("No conversion yet.");
  });

  it("shows an error when convert is clicked before loading a workbook", () => {
    bootMain();

    document.getElementById("convertBtn").click();

    expect(document.getElementById("errorAlert").dataset.message).toBe("Load an xlsx file first.");
  });

  it("loads a workbook from the file input and passes UI options to conversion", async () => {
    const api = bootMain();
    document.getElementById("includeShapeDetailsEnabled").checked = false;
    document.getElementById("outputModeSelect").value = "both";
    document.getElementById("formattingModeSelect").value = "github";
    document.getElementById("tableDetectionModeSelect").value = "border";

    const fileInput = document.getElementById("xlsxFileInput");
    const file = {
      name: "sample.xlsx",
      arrayBuffer: async () => new ArrayBuffer(8)
    };
    Object.defineProperty(fileInput, "files", {
      configurable: true,
      get: () => [file]
    });

    fileInput.dispatchEvent(new Event("change"));
    await flushAsyncWork();

    expect(api.parseWorkbook).toHaveBeenCalledWith(expect.any(ArrayBuffer), "sample.xlsx");
    expect(api.convertWorkbookToMarkdownFiles).toHaveBeenCalledWith(
      expect.objectContaining({ name: "book.xlsx" }),
      expect.objectContaining({
        includeShapeDetails: false,
        outputMode: "both",
        formattingMode: "github",
        tableDetectionMode: "border",
        treatFirstRowAsHeader: true,
        trimText: true,
        removeEmptyRows: true,
        removeEmptyColumns: true
      })
    );
    expect(document.getElementById("downloadBtn").disabled).toBe(false);
    expect(document.getElementById("exportZipBtn").disabled).toBe(false);
    expect(document.getElementById("markdownPreview").dataset.rendered).toContain("# Sheet1");
  });

  it("disables shift_jis in browser-only runtime", () => {
    bootMain({}, { disableNodeRequire: true });

    const encodingSelect = document.getElementById("encodingSelect");
    const shiftJisOption = Array.from(encodingSelect.options).find((option) => option.value === "shift_jis");

    expect(shiftJisOption.disabled).toBe(true);
    expect(shiftJisOption.text).toContain("CLI only");
    expect(document.getElementById("encodingNotice").textContent).toContain("`utf-8`");
  });
});
