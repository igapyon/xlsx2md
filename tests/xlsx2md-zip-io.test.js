// @vitest-environment jsdom

import { Blob as NodeBlob } from "node:buffer";
import { readFileSync } from "node:fs";
import path from "node:path";
import { DecompressionStream as NodeDecompressionStream } from "node:stream/web";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";
import { loadModuleRegistry } from "./helpers/module-registry.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

if (typeof globalThis.Blob === "undefined" || typeof globalThis.Blob.prototype?.stream !== "function") {
  globalThis.Blob = NodeBlob;
}
globalThis.DecompressionStream ??= NodeDecompressionStream;

const zipIoCode = readFileSync(
  path.resolve(__dirname, "../src/xlsx2md/js/zip-io.js"),
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
});
