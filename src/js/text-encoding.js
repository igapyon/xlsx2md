/*
 * Copyright 2026 Toshiki Iga
 * SPDX-License-Identifier: Apache-2.0
 */
(() => {
    const moduleRegistry = getXlsx2mdModuleRegistry();
    const utf8Encoder = new TextEncoder();
    const nodeRequire = (() => {
        const candidate = globalThis.__xlsx2mdNodeRequire;
        return typeof candidate === "function" ? candidate : null;
    })();
    function normalizeEncoding(value) {
        const normalized = String(value || "utf-8").toLowerCase();
        if (normalized === "utf-8" ||
            normalized === "shift_jis" ||
            normalized === "utf-16le" ||
            normalized === "utf-16be" ||
            normalized === "utf-32le" ||
            normalized === "utf-32be") {
            return normalized;
        }
        throw new Error(`Unsupported encoding: ${String(value)}`);
    }
    function normalizeBom(value) {
        const normalized = String(value || "off").toLowerCase();
        if (normalized === "off" || normalized === "on") {
            return normalized;
        }
        throw new Error(`Unsupported BOM mode: ${String(value)}`);
    }
    function concatBytes(parts) {
        const totalLength = parts.reduce((sum, part) => sum + part.length, 0);
        const result = new Uint8Array(totalLength);
        let offset = 0;
        for (const part of parts) {
            result.set(part, offset);
            offset += part.length;
        }
        return result;
    }
    function getBomBytes(encoding) {
        if (encoding === "utf-8")
            return new Uint8Array([0xef, 0xbb, 0xbf]);
        if (encoding === "utf-16le")
            return new Uint8Array([0xff, 0xfe]);
        if (encoding === "utf-16be")
            return new Uint8Array([0xfe, 0xff]);
        if (encoding === "utf-32le")
            return new Uint8Array([0xff, 0xfe, 0x00, 0x00]);
        if (encoding === "utf-32be")
            return new Uint8Array([0x00, 0x00, 0xfe, 0xff]);
        throw new Error(`Encoding does not support BOM: ${encoding}`);
    }
    function encodeUtf16(text, littleEndian) {
        const result = new Uint8Array(text.length * 2);
        for (let index = 0; index < text.length; index += 1) {
            const codeUnit = text.charCodeAt(index);
            const offset = index * 2;
            if (littleEndian) {
                result[offset] = codeUnit & 0xff;
                result[offset + 1] = codeUnit >>> 8;
            }
            else {
                result[offset] = codeUnit >>> 8;
                result[offset + 1] = codeUnit & 0xff;
            }
        }
        return result;
    }
    function encodeUtf32(text, littleEndian) {
        const codePoints = [];
        for (let index = 0; index < text.length; index += 1) {
            const first = text.charCodeAt(index);
            if (first >= 0xd800 && first <= 0xdbff && index + 1 < text.length) {
                const second = text.charCodeAt(index + 1);
                if (second >= 0xdc00 && second <= 0xdfff) {
                    codePoints.push(((first - 0xd800) << 10) + (second - 0xdc00) + 0x10000);
                    index += 1;
                    continue;
                }
            }
            codePoints.push(first);
        }
        const result = new Uint8Array(codePoints.length * 4);
        codePoints.forEach((codePoint, index) => {
            const offset = index * 4;
            if (littleEndian) {
                result[offset] = codePoint & 0xff;
                result[offset + 1] = (codePoint >>> 8) & 0xff;
                result[offset + 2] = (codePoint >>> 16) & 0xff;
                result[offset + 3] = (codePoint >>> 24) & 0xff;
            }
            else {
                result[offset] = (codePoint >>> 24) & 0xff;
                result[offset + 1] = (codePoint >>> 16) & 0xff;
                result[offset + 2] = (codePoint >>> 8) & 0xff;
                result[offset + 3] = codePoint & 0xff;
            }
        });
        return result;
    }
    function encodeText(text, options = {}) {
        const encoding = normalizeEncoding(options.encoding);
        const bom = normalizeBom(options.bom);
        if (encoding === "shift_jis") {
            if (bom === "on") {
                throw new Error("BOM cannot be enabled for shift_jis.");
            }
            if (!nodeRequire) {
                throw new Error("Shift_JIS encoding is not available in this runtime.");
            }
            const iconvLite = nodeRequire("iconv-lite");
            return Uint8Array.from(iconvLite.encode(text, "shift_jis"));
        }
        const body = encoding === "utf-8"
            ? utf8Encoder.encode(text)
            : encoding === "utf-16le"
                ? encodeUtf16(text, true)
                : encoding === "utf-16be"
                    ? encodeUtf16(text, false)
                    : encoding === "utf-32le"
                        ? encodeUtf32(text, true)
                        : encodeUtf32(text, false);
        if (bom === "off") {
            return body;
        }
        return concatBytes([getBomBytes(encoding), body]);
    }
    function isEncodingAvailable(value) {
        const encoding = normalizeEncoding(value);
        if (encoding === "shift_jis") {
            return !!nodeRequire;
        }
        return true;
    }
    function createTextMimeType(options = {}) {
        return `text/markdown;charset=${normalizeEncoding(options.encoding)}`;
    }
    const textEncodingApi = {
        normalizeEncoding,
        normalizeBom,
        getBomBytes,
        isEncodingAvailable,
        encodeText,
        createTextMimeType
    };
    moduleRegistry.registerModule("textEncoding", textEncodingApi);
})();
