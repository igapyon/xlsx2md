(() => {
  type ZipEntryRecord = {
    name: string;
    compressionMethod: number;
    compressedSize: number;
    uncompressedSize: number;
    localHeaderOffset: number;
  };

  type ExportEntry = {
    name: string;
    data: Uint8Array;
  };

  const textDecoder = new TextDecoder("utf-8");
  const textEncoder = new TextEncoder();
  const crcTable = buildCrc32Table();

  function buildCrc32Table(): Uint32Array {
    const table = new Uint32Array(256);
    for (let i = 0; i < 256; i += 1) {
      let value = i;
      for (let bit = 0; bit < 8; bit += 1) {
        value = (value & 1) === 1 ? (0xedb88320 ^ (value >>> 1)) : (value >>> 1);
      }
      table[i] = value >>> 0;
    }
    return table;
  }

  function crc32(bytes: Uint8Array): number {
    let crc = 0xffffffff;
    for (const byte of bytes) {
      crc = crcTable[(crc ^ byte) & 0xff] ^ (crc >>> 8);
    }
    return (crc ^ 0xffffffff) >>> 0;
  }

  function decodeXmlText(bytes: Uint8Array): string {
    return textDecoder.decode(bytes);
  }

  function readUint16LE(view: DataView, offset: number): number {
    return view.getUint16(offset, true);
  }

  function readUint32LE(view: DataView, offset: number): number {
    return view.getUint32(offset, true);
  }

  async function inflateRaw(data: Uint8Array): Promise<Uint8Array> {
    if (typeof DecompressionStream === "function") {
      const stream = new Blob([data]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
      const buffer = await new Response(stream).arrayBuffer();
      return new Uint8Array(buffer);
    }
    throw new Error("This environment does not support ZIP deflate decompression.");
  }

  async function unzipEntries(arrayBuffer: ArrayBuffer): Promise<Map<string, Uint8Array>> {
    const view = new DataView(arrayBuffer);
    let eocdOffset = -1;
    for (let offset = view.byteLength - 22; offset >= Math.max(0, view.byteLength - 0x10000 - 22); offset -= 1) {
      if (readUint32LE(view, offset) === 0x06054b50) {
        eocdOffset = offset;
        break;
      }
    }
    if (eocdOffset < 0) {
      throw new Error("ZIP end-of-central-directory record was not found.");
    }

    const centralDirectorySize = readUint32LE(view, eocdOffset + 12);
    const centralDirectoryOffset = readUint32LE(view, eocdOffset + 16);
    const endOffset = centralDirectoryOffset + centralDirectorySize;
    const entries: ZipEntryRecord[] = [];
    let cursor = centralDirectoryOffset;

    while (cursor < endOffset) {
      if (readUint32LE(view, cursor) !== 0x02014b50) {
        throw new Error("ZIP central directory format is invalid.");
      }
      const compressionMethod = readUint16LE(view, cursor + 10);
      const compressedSize = readUint32LE(view, cursor + 20);
      const uncompressedSize = readUint32LE(view, cursor + 24);
      const fileNameLength = readUint16LE(view, cursor + 28);
      const extraFieldLength = readUint16LE(view, cursor + 30);
      const fileCommentLength = readUint16LE(view, cursor + 32);
      const localHeaderOffset = readUint32LE(view, cursor + 42);
      const fileNameBytes = new Uint8Array(arrayBuffer, cursor + 46, fileNameLength);
      const name = decodeXmlText(fileNameBytes);
      entries.push({
        name,
        compressionMethod,
        compressedSize,
        uncompressedSize,
        localHeaderOffset
      });
      cursor += 46 + fileNameLength + extraFieldLength + fileCommentLength;
    }

    const files = new Map<string, Uint8Array>();
    for (const entry of entries) {
      const localOffset = entry.localHeaderOffset;
      if (readUint32LE(view, localOffset) !== 0x04034b50) {
        throw new Error(`ZIP local header is invalid: ${entry.name}`);
      }
      const fileNameLength = readUint16LE(view, localOffset + 26);
      const extraFieldLength = readUint16LE(view, localOffset + 28);
      const dataOffset = localOffset + 30 + fileNameLength + extraFieldLength;
      const compressedData = new Uint8Array(arrayBuffer, dataOffset, entry.compressedSize);

      let fileData: Uint8Array;
      if (entry.compressionMethod === 0) {
        fileData = new Uint8Array(compressedData);
      } else if (entry.compressionMethod === 8) {
        fileData = await inflateRaw(compressedData);
      } else {
        throw new Error(`Unsupported compression method: ${entry.name} (method=${entry.compressionMethod})`);
      }

      files.set(entry.name, fileData);
    }

    return files;
  }

  function createStoredZip(entries: ExportEntry[]): Uint8Array {
    const localChunks: Uint8Array[] = [];
    const centralChunks: Uint8Array[] = [];
    let offset = 0;

    for (const entry of entries) {
      const nameBytes = textEncoder.encode(entry.name);
      const dataBytes = entry.data;
      const entryCrc32 = crc32(dataBytes);

      const localHeader = new Uint8Array(30 + nameBytes.length);
      const localView = new DataView(localHeader.buffer);
      localView.setUint32(0, 0x04034b50, true);
      localView.setUint16(4, 20, true);
      localView.setUint16(6, 0, true);
      localView.setUint16(8, 0, true);
      localView.setUint16(10, 0, true);
      localView.setUint16(12, 0, true);
      localView.setUint32(14, entryCrc32, true);
      localView.setUint32(18, dataBytes.length, true);
      localView.setUint32(22, dataBytes.length, true);
      localView.setUint16(26, nameBytes.length, true);
      localView.setUint16(28, 0, true);
      localHeader.set(nameBytes, 30);
      localChunks.push(localHeader, dataBytes);

      const centralHeader = new Uint8Array(46 + nameBytes.length);
      const centralView = new DataView(centralHeader.buffer);
      centralView.setUint32(0, 0x02014b50, true);
      centralView.setUint16(4, 20, true);
      centralView.setUint16(6, 20, true);
      centralView.setUint16(8, 0, true);
      centralView.setUint16(10, 0, true);
      centralView.setUint16(12, 0, true);
      centralView.setUint16(14, 0, true);
      centralView.setUint32(16, entryCrc32, true);
      centralView.setUint32(20, dataBytes.length, true);
      centralView.setUint32(24, dataBytes.length, true);
      centralView.setUint16(28, nameBytes.length, true);
      centralView.setUint16(30, 0, true);
      centralView.setUint16(32, 0, true);
      centralView.setUint16(34, 0, true);
      centralView.setUint16(36, 0, true);
      centralView.setUint32(38, 0, true);
      centralView.setUint32(42, offset, true);
      centralHeader.set(nameBytes, 46);
      centralChunks.push(centralHeader);

      offset += localHeader.length + dataBytes.length;
    }

    const centralDirectoryStart = offset;
    const centralDirectorySize = centralChunks.reduce((sum, chunk) => sum + chunk.length, 0);
    const eocd = new Uint8Array(22);
    const eocdView = new DataView(eocd.buffer);
    eocdView.setUint32(0, 0x06054b50, true);
    eocdView.setUint16(4, 0, true);
    eocdView.setUint16(6, 0, true);
    eocdView.setUint16(8, entries.length, true);
    eocdView.setUint16(10, entries.length, true);
    eocdView.setUint32(12, centralDirectorySize, true);
    eocdView.setUint32(16, centralDirectoryStart, true);
    eocdView.setUint16(20, 0, true);

    const totalLength = localChunks.reduce((sum, chunk) => sum + chunk.length, 0) + centralDirectorySize + eocd.length;
    const output = new Uint8Array(totalLength);
    let cursor = 0;
    for (const chunk of localChunks) {
      output.set(chunk, cursor);
      cursor += chunk.length;
    }
    for (const chunk of centralChunks) {
      output.set(chunk, cursor);
      cursor += chunk.length;
    }
    output.set(eocd, cursor);
    return output;
  }

  (globalThis as typeof globalThis & {
    __xlsx2mdZipIo?: {
      unzipEntries: typeof unzipEntries;
      createStoredZip: typeof createStoredZip;
    };
  }).__xlsx2mdZipIo = {
    unzipEntries,
    createStoredZip
  };
})();
