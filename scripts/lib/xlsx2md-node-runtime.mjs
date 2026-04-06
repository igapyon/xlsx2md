import { Blob as NodeBlob } from "node:buffer";
import fs from "node:fs";
import { createRequire } from "node:module";
import path from "node:path";
import { DecompressionStream as NodeDecompressionStream } from "node:stream/web";
import { fileURLToPath } from "node:url";

import {
  DOMParser as XmldomParser,
  XMLSerializer as XmldomSerializer,
  Node as XmldomNode,
  Document as XmldomDocument,
  Element as XmldomElement
} from "@xmldom/xmldom";

import { XLSX2MD_CORE_JS_ORDER } from "./xlsx2md-module-order.mjs";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const DEFAULT_ROOT_DIR = path.resolve(__dirname, "../..");
const nodeRequire = createRequire(import.meta.url);

let cachedApi = null;
let cachedRootDir = null;

export function installNodeDomGlobals() {
  if (typeof globalThis.DOMParser !== "function") {
    globalThis.DOMParser = XmldomParser;
    globalThis.Node = XmldomNode;
    globalThis.Document = XmldomDocument;
    globalThis.Element = XmldomElement;
    globalThis.XMLSerializer ??= XmldomSerializer;
  }

  if (typeof globalThis.Blob === "undefined" || typeof globalThis.Blob.prototype?.stream !== "function") {
    globalThis.Blob = NodeBlob;
  }
  globalThis.DecompressionStream ??= NodeDecompressionStream;
  globalThis.__xlsx2mdNodeRequire ??= nodeRequire;
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
