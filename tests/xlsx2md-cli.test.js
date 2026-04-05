// @vitest-environment node

import { mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

const execFileAsync = promisify(execFile);
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

describe("xlsx2md cli", () => {
  it("writes combined markdown from a workbook", async () => {
    const workspace = mkdtempSync(path.join(tmpdir(), "xlsx2md-cli-"));
    const fixturePath = path.resolve(__dirname, "./fixtures/xlsx2md-basic-sample01.xlsx");
    const outputPath = path.join(workspace, "result.md");

    try {
      await execFileAsync(process.execPath, [
        path.resolve(__dirname, "../scripts/xlsx2md-cli.mjs"),
        fixturePath,
        "--out",
        outputPath
      ], {
        cwd: path.resolve(__dirname, "..")
      });

      const outputText = readFileSync(outputPath, "utf8");
      expect(outputText).toContain("# ");
      expect(outputText).toContain("<!-- ");
    } finally {
      rmSync(workspace, { recursive: true, force: true });
    }
  });

  it("writes a zip export from a workbook", async () => {
    const workspace = mkdtempSync(path.join(tmpdir(), "xlsx2md-cli-"));
    const fixturePath = path.resolve(__dirname, "./fixtures/xlsx2md-basic-sample01.xlsx");
    const outputPath = path.join(workspace, "result.zip");

    try {
      await execFileAsync(process.execPath, [
        path.resolve(__dirname, "../scripts/xlsx2md-cli.mjs"),
        fixturePath,
        "--zip",
        outputPath
      ], {
        cwd: path.resolve(__dirname, "..")
      });

      const outputBytes = readFileSync(outputPath);
      expect(outputBytes.length).toBeGreaterThan(0);
    } finally {
      rmSync(workspace, { recursive: true, force: true });
    }
  });

  it("reports the input workbook when the command fails", async () => {
    await expect(execFileAsync(process.execPath, [
      path.resolve(__dirname, "../scripts/xlsx2md-cli.mjs"),
      path.resolve(__dirname, "./fixtures/does-not-exist.xlsx")
    ], {
      cwd: path.resolve(__dirname, "..")
    })).rejects.toMatchObject({
      stderr: expect.stringContaining("[does-not-exist.xlsx] read failed:")
    });
  });

  it("prints help and exits successfully", async () => {
    const result = await execFileAsync(process.execPath, [
      path.resolve(__dirname, "../scripts/xlsx2md-cli.mjs"),
      "--help"
    ], {
      cwd: path.resolve(__dirname, "..")
    });

    expect(result.stdout).toContain("Usage:");
    expect(result.stdout).toContain("--shape-details");
    expect(result.stdout).toContain("--include-shape-details");
    expect(result.stdout).toContain("--encoding");
    expect(result.stdout).toContain("--bom");
    expect(result.stdout).toContain("--formatting-mode");
    expect(result.stdout).toContain("--table-detection-mode");
    expect(result.stdout).toContain("Exit codes:");
  });

  it("fails for an unknown option", async () => {
    await expect(execFileAsync(process.execPath, [
      path.resolve(__dirname, "../scripts/xlsx2md-cli.mjs"),
      path.resolve(__dirname, "./fixtures/xlsx2md-basic-sample01.xlsx"),
      "--unknown-option"
    ], {
      cwd: path.resolve(__dirname, "..")
    })).rejects.toMatchObject({
      stderr: expect.stringContaining("Unknown option: --unknown-option")
    });
  });

  it("prints a summary when requested", async () => {
    const workspace = mkdtempSync(path.join(tmpdir(), "xlsx2md-cli-"));

    try {
      const result = await execFileAsync(process.execPath, [
        path.resolve(__dirname, "../scripts/xlsx2md-cli.mjs"),
        path.resolve(__dirname, "./fixtures/xlsx2md-basic-sample01.xlsx"),
        "--summary"
      ], {
        cwd: workspace
      });

      expect(result.stdout).toContain("[workbook] xlsx2md-basic-sample01.xlsx");
      expect(result.stdout).toContain("Output file:");
    } finally {
      rmSync(workspace, { recursive: true, force: true });
    }
  });

  it("accepts output mode and shape detail options", async () => {
    const workspace = mkdtempSync(path.join(tmpdir(), "xlsx2md-cli-"));
    const fixturePath = path.resolve(__dirname, "./fixtures/shape/shape-basic-sample01.xlsx");
    const outputPath = path.join(workspace, "shape.md");

    try {
      await execFileAsync(process.execPath, [
        path.resolve(__dirname, "../scripts/xlsx2md-cli.mjs"),
        fixturePath,
        "--out",
        outputPath,
        "--output-mode",
        "both",
        "--formatting-mode",
        "github",
        "--table-detection-mode",
        "border",
        "--shape-details",
        "include"
      ], {
        cwd: path.resolve(__dirname, "..")
      });

      const outputText = readFileSync(outputPath, "utf8");
      expect(outputText).toContain("# ");
    } finally {
      rmSync(workspace, { recursive: true, force: true });
    }
  });

  it("writes UTF-16BE markdown with BOM when requested", async () => {
    const workspace = mkdtempSync(path.join(tmpdir(), "xlsx2md-cli-"));
    const fixturePath = path.resolve(__dirname, "./fixtures/xlsx2md-basic-sample01.xlsx");
    const outputPath = path.join(workspace, "utf16be.md");

    try {
      await execFileAsync(process.execPath, [
        path.resolve(__dirname, "../scripts/xlsx2md-cli.mjs"),
        fixturePath,
        "--out",
        outputPath,
        "--encoding",
        "utf-16be",
        "--bom",
        "on"
      ], {
        cwd: path.resolve(__dirname, "..")
      });

      const outputBytes = readFileSync(outputPath);
      expect(Array.from(outputBytes.slice(0, 2))).toEqual([0xfe, 0xff]);
    } finally {
      rmSync(workspace, { recursive: true, force: true });
    }
  });

  it("writes Shift_JIS markdown when requested", async () => {
    const workspace = mkdtempSync(path.join(tmpdir(), "xlsx2md-cli-"));
    const fixturePath = path.resolve(__dirname, "./fixtures/xlsx2md-basic-sample01.xlsx");
    const outputPath = path.join(workspace, "sjis.md");

    try {
      await execFileAsync(process.execPath, [
        path.resolve(__dirname, "../scripts/xlsx2md-cli.mjs"),
        fixturePath,
        "--out",
        outputPath,
        "--encoding",
        "shift_jis"
      ], {
        cwd: path.resolve(__dirname, "..")
      });

      const outputBytes = readFileSync(outputPath);
      const outputText = new TextDecoder("shift_jis").decode(outputBytes);
      expect(outputText).toContain("# ");
      expect(outputText).toContain("<!-- ");
    } finally {
      rmSync(workspace, { recursive: true, force: true });
    }
  });

  it("keeps --include-shape-details as a compatibility alias", async () => {
    const workspace = mkdtempSync(path.join(tmpdir(), "xlsx2md-cli-"));
    const fixturePath = path.resolve(__dirname, "./fixtures/shape/shape-basic-sample01.xlsx");
    const outputPath = path.join(workspace, "shape-alias.md");

    try {
      await execFileAsync(process.execPath, [
        path.resolve(__dirname, "../scripts/xlsx2md-cli.mjs"),
        fixturePath,
        "--out",
        outputPath,
        "--include-shape-details"
      ], {
        cwd: path.resolve(__dirname, "..")
      });

      const outputText = readFileSync(outputPath, "utf8");
      expect(outputText).toContain("# ");
    } finally {
      rmSync(workspace, { recursive: true, force: true });
    }
  }, 15000);

  it("fails for an invalid output mode", async () => {
    await expect(execFileAsync(process.execPath, [
      path.resolve(__dirname, "../scripts/xlsx2md-cli.mjs"),
      path.resolve(__dirname, "./fixtures/xlsx2md-basic-sample01.xlsx"),
      "--output-mode",
      "invalid"
    ], {
      cwd: path.resolve(__dirname, "..")
    })).rejects.toMatchObject({
      stderr: expect.stringContaining("Invalid output mode: invalid")
    });
  });

  it("fails for an invalid formatting mode", async () => {
    await expect(execFileAsync(process.execPath, [
      path.resolve(__dirname, "../scripts/xlsx2md-cli.mjs"),
      path.resolve(__dirname, "./fixtures/xlsx2md-basic-sample01.xlsx"),
      "--formatting-mode",
      "invalid"
    ], {
      cwd: path.resolve(__dirname, "..")
    })).rejects.toMatchObject({
      stderr: expect.stringContaining("Invalid formatting mode: invalid")
    });
  });

  it("fails for an invalid shape details mode", async () => {
    await expect(execFileAsync(process.execPath, [
      path.resolve(__dirname, "../scripts/xlsx2md-cli.mjs"),
      path.resolve(__dirname, "./fixtures/xlsx2md-basic-sample01.xlsx"),
      "--shape-details",
      "invalid"
    ], {
      cwd: path.resolve(__dirname, "..")
    })).rejects.toMatchObject({
      stderr: expect.stringContaining("Invalid shape details mode: invalid")
    });
  });

  it("fails for an invalid table detection mode", async () => {
    await expect(execFileAsync(process.execPath, [
      path.resolve(__dirname, "../scripts/xlsx2md-cli.mjs"),
      path.resolve(__dirname, "./fixtures/xlsx2md-basic-sample01.xlsx"),
      "--table-detection-mode",
      "invalid"
    ], {
      cwd: path.resolve(__dirname, "..")
    })).rejects.toMatchObject({
      stderr: expect.stringContaining("Invalid table detection mode: invalid")
    });
  });

  it("fails for an invalid encoding", async () => {
    await expect(execFileAsync(process.execPath, [
      path.resolve(__dirname, "../scripts/xlsx2md-cli.mjs"),
      path.resolve(__dirname, "./fixtures/xlsx2md-basic-sample01.xlsx"),
      "--encoding",
      "invalid"
    ], {
      cwd: path.resolve(__dirname, "..")
    })).rejects.toMatchObject({
      stderr: expect.stringContaining("Invalid encoding: invalid")
    });
  });

  it("fails for an invalid BOM mode", async () => {
    await expect(execFileAsync(process.execPath, [
      path.resolve(__dirname, "../scripts/xlsx2md-cli.mjs"),
      path.resolve(__dirname, "./fixtures/xlsx2md-basic-sample01.xlsx"),
      "--bom",
      "invalid"
    ], {
      cwd: path.resolve(__dirname, "..")
    })).rejects.toMatchObject({
      stderr: expect.stringContaining("Invalid BOM mode: invalid")
    });
  });

  it("fails when BOM is enabled for shift_jis", async () => {
    await expect(execFileAsync(process.execPath, [
      path.resolve(__dirname, "../scripts/xlsx2md-cli.mjs"),
      path.resolve(__dirname, "./fixtures/xlsx2md-basic-sample01.xlsx"),
      "--encoding",
      "shift_jis",
      "--bom",
      "on"
    ], {
      cwd: path.resolve(__dirname, "..")
    })).rejects.toMatchObject({
      stderr: expect.stringContaining("BOM cannot be enabled for shift_jis.")
    });
  });
});
