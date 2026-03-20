import { Blob as NodeBlob } from "node:buffer";
import fs from "node:fs";
import path from "node:path";
import { DecompressionStream as NodeDecompressionStream } from "node:stream/web";
import { fileURLToPath } from "node:url";

import { JSDOM } from "jsdom";

import { XLSX2MD_CORE_JS_ORDER } from "./xlsx2md-module-order.mjs";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const DEFAULT_ROOT_DIR = path.resolve(__dirname, "../..");

let cachedApi = null;
let cachedRootDir = null;

export function installNodeDomGlobals() {
  if (typeof globalThis.DOMParser !== "function") {
    const dom = new JSDOM("<!doctype html><html><body></body></html>");
    globalThis.window ??= dom.window;
    globalThis.document ??= dom.window.document;
    globalThis.DOMParser = dom.window.DOMParser;
    globalThis.Node = dom.window.Node;
    globalThis.Document = dom.window.Document;
    globalThis.Element = dom.window.Element;
    globalThis.ParentNode = dom.window.ParentNode;
    globalThis.XMLSerializer ??= dom.window.XMLSerializer;
  }

  if (typeof globalThis.Blob === "undefined" || typeof globalThis.Blob.prototype?.stream !== "function") {
    globalThis.Blob = NodeBlob;
  }
  globalThis.DecompressionStream ??= NodeDecompressionStream;
}

export function loadXlsx2mdNodeApi(options = {}) {
  const rootDir = path.resolve(options.rootDir || DEFAULT_ROOT_DIR);
  if (cachedApi && cachedRootDir === rootDir) {
    return cachedApi;
  }

  installNodeDomGlobals();

  for (const relPath of XLSX2MD_CORE_JS_ORDER) {
    const absPath = path.resolve(rootDir, relPath);
    const source = fs.readFileSync(absPath, "utf8");
    new Function(source)();
  }

  const api = globalThis.__xlsx2mdModuleRegistry?.getModule("xlsx2md");
  if (!api) {
    throw new Error("xlsx2md core API failed to initialize.");
  }

  cachedApi = api;
  cachedRootDir = rootDir;
  return api;
}
