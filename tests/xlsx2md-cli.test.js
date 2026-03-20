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
});
