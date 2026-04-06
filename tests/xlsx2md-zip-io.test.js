// @vitest-environment jsdom

import { Blob as NodeBlob } from "node:buffer";
import { createRequire } from "node:module";
import { readFileSync } from "node:fs";
import path from "node:path";
import { DecompressionStream as NodeDecompressionStream } from "node:stream/web";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";
import { loadModuleRegistry } from "./helpers/module-registry.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const nodeRequire = createRequire(import.meta.url);

if (typeof globalThis.Blob === "undefined" || typeof globalThis.Blob.prototype?.stream !== "function") {
  globalThis.Blob = NodeBlob;
}
globalThis.DecompressionStream ??= NodeDecompressionStream;
globalThis.__xlsx2mdNodeRequire ??= nodeRequire;

const zipIoCode = readFileSync(
  path.resolve(__dirname, "../src/js/zip-io.js"),
  "utf8"
);

function bootZipIo() {
  document.body.innerHTML = "";
  loadModuleRegistry(__dirname);
  new Function(zipIoCode)();
  return globalThis.__xlsx2mdModuleRegistry.getModule("zipIo");
}

describe("xlsx2md zip io", () => {
  it("round-trips stored zip entries", async () => {
    const api = bootZipIo();
    const encoder = new TextEncoder();
    const zipBytes = api.createStoredZip([
      { name: "output/test.md", data: encoder.encode("# Test\n") },
      { name: "output/assets/icon.txt", data: encoder.encode("asset") }
    ]);

    const extracted = await api.unzipEntries(zipBytes.buffer.slice(zipBytes.byteOffset, zipBytes.byteOffset + zipBytes.byteLength));

    expect(Array.from(extracted.keys()).sort()).toEqual(["output/assets/icon.txt", "output/test.md"]);
    expect(new TextDecoder().decode(extracted.get("output/test.md"))).toBe("# Test\n");
    expect(new TextDecoder().decode(extracted.get("output/assets/icon.txt"))).toBe("asset");
  });

  it("supports empty file payloads", async () => {
    const api = bootZipIo();
    const zipBytes = api.createStoredZip([
      { name: "empty.txt", data: new Uint8Array([]) }
    ]);

    const extracted = await api.unzipEntries(zipBytes.buffer.slice(zipBytes.byteOffset, zipBytes.byteOffset + zipBytes.byteLength));

    expect(extracted.has("empty.txt")).toBe(true);
    expect(extracted.get("empty.txt")).toEqual(new Uint8Array([]));
  });

  it("writes a fixed reproducible ZIP entry timestamp", () => {
    const api = bootZipIo();
    const encoder = new TextEncoder();
    const zipBytes = api.createStoredZip([
      { name: "output/test.md", data: encoder.encode("# Test\n") }
    ]);
    const view = new DataView(zipBytes.buffer, zipBytes.byteOffset, zipBytes.byteLength);

    expect(view.getUint32(0, true)).toBe(0x04034b50);
    expect(view.getUint16(10, true)).toBe(api.fixedZipEntryTimestamp.dosTime);
    expect(view.getUint16(12, true)).toBe(api.fixedZipEntryTimestamp.dosDate);

    const localNameLength = view.getUint16(26, true);
    const centralOffset = 30 + localNameLength + encoder.encode("# Test\n").length;
    expect(view.getUint32(centralOffset, true)).toBe(0x02014b50);
    expect(view.getUint16(centralOffset + 12, true)).toBe(api.fixedZipEntryTimestamp.dosTime);
    expect(view.getUint16(centralOffset + 14, true)).toBe(api.fixedZipEntryTimestamp.dosDate);
  });

  it("does not mark ASCII-only file names with the UTF-8 flag", () => {
    const api = bootZipIo();
    const encoder = new TextEncoder();
    const zipBytes = api.createStoredZip([
      { name: "output/test.md", data: encoder.encode("# Test\n") }
    ]);
    const view = new DataView(zipBytes.buffer, zipBytes.byteOffset, zipBytes.byteLength);

    expect(view.getUint16(6, true)).toBe(0);

    const localNameLength = view.getUint16(26, true);
    const centralOffset = 30 + localNameLength + encoder.encode("# Test\n").length;
    expect(view.getUint16(centralOffset + 8, true)).toBe(0);
  });

  it("marks UTF-8 file names so non-ASCII entries unzip correctly", async () => {
    const api = bootZipIo();
    const encoder = new TextEncoder();
    const zipBytes = api.createStoredZip([
      { name: "output/日本語.md", data: encoder.encode("# 日本語\n") }
    ]);
    const view = new DataView(zipBytes.buffer, zipBytes.byteOffset, zipBytes.byteLength);

    expect(view.getUint16(6, true)).toBe(0x0800);

    const localNameLength = view.getUint16(26, true);
    const centralOffset = 30 + localNameLength + encoder.encode("# 日本語\n").length;
    expect(view.getUint16(centralOffset + 8, true)).toBe(0x0800);

    const extracted = await api.unzipEntries(
      zipBytes.buffer.slice(zipBytes.byteOffset, zipBytes.byteOffset + zipBytes.byteLength)
    );
    expect(Array.from(extracted.keys())).toEqual(["output/日本語.md"]);
    expect(new TextDecoder().decode(extracted.get("output/日本語.md"))).toBe("# 日本語\n");
  });

  it("throws for invalid zip input", async () => {
    const api = bootZipIo();

    await expect(api.unzipEntries(new Uint8Array([1, 2, 3, 4]).buffer)).rejects.toThrow(
      "ZIP end-of-central-directory record was not found."
    );
  });

  it("falls back to node zlib inflateRawSync when deflate-raw streams are unavailable", async () => {
    const api = bootZipIo();
    const encoder = new TextEncoder();
    const zlib = nodeRequire("node:zlib");
    const fileName = "output/test.txt";
    const fileNameBytes = encoder.encode(fileName);
    const original = encoder.encode("fallback");
    const compressed = Uint8Array.from(zlib.deflateRawSync(original));
    const crc = 0xa87e4381;

    const localHeader = new Uint8Array(30 + fileNameBytes.length);
    const localView = new DataView(localHeader.buffer);
    localView.setUint32(0, 0x04034b50, true);
    localView.setUint16(4, 20, true);
    localView.setUint16(6, 0, true);
    localView.setUint16(8, 8, true);
    localView.setUint16(10, 0, true);
    localView.setUint16(12, 0, true);
    localView.setUint32(14, crc, true);
    localView.setUint32(18, compressed.length, true);
    localView.setUint32(22, original.length, true);
    localView.setUint16(26, fileNameBytes.length, true);
    localView.setUint16(28, 0, true);
    localHeader.set(fileNameBytes, 30);

    const centralHeader = new Uint8Array(46 + fileNameBytes.length);
    const centralView = new DataView(centralHeader.buffer);
    centralView.setUint32(0, 0x02014b50, true);
    centralView.setUint16(4, 20, true);
    centralView.setUint16(6, 20, true);
    centralView.setUint16(8, 0, true);
    centralView.setUint16(10, 8, true);
    centralView.setUint16(12, 0, true);
    centralView.setUint16(14, 0, true);
    centralView.setUint32(16, crc, true);
    centralView.setUint32(20, compressed.length, true);
    centralView.setUint32(24, original.length, true);
    centralView.setUint16(28, fileNameBytes.length, true);
    centralView.setUint16(30, 0, true);
    centralView.setUint16(32, 0, true);
    centralView.setUint16(34, 0, true);
    centralView.setUint16(36, 0, true);
    centralView.setUint32(38, 0, true);
    centralView.setUint32(42, 0, true);
    centralHeader.set(fileNameBytes, 46);

    const eocd = new Uint8Array(22);
    const eocdView = new DataView(eocd.buffer);
    eocdView.setUint32(0, 0x06054b50, true);
    eocdView.setUint16(8, 1, true);
    eocdView.setUint16(10, 1, true);
    eocdView.setUint32(12, centralHeader.length, true);
    eocdView.setUint32(16, localHeader.length + compressed.length, true);

    const zipBytes = new Uint8Array(localHeader.length + compressed.length + centralHeader.length + eocd.length);
    zipBytes.set(localHeader, 0);
    zipBytes.set(compressed, localHeader.length);
    zipBytes.set(centralHeader, localHeader.length + compressed.length);
    zipBytes.set(eocd, localHeader.length + compressed.length + centralHeader.length);

    const originalDecompressionStream = globalThis.DecompressionStream;
    globalThis.DecompressionStream = class UnsupportedDecompressionStream {
      constructor(format) {
        throw new Error(`unsupported: ${format}`);
      }
    };

    try {
      const extracted = await api.unzipEntries(
        zipBytes.buffer.slice(zipBytes.byteOffset, zipBytes.byteOffset + zipBytes.byteLength)
      );
      expect(new TextDecoder().decode(extracted.get(fileName))).toBe("fallback");
    } finally {
      globalThis.DecompressionStream = originalDecompressionStream;
    }
  });
});
