import { execFileSync } from "node:child_process";
import { readdirSync, statSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

function collectFixtureXlsxFiles(dirPath) {
  const entries = readdirSync(dirPath).sort();
  const files = [];
  for (const entry of entries) {
    const fullPath = path.join(dirPath, entry);
    const stat = statSync(fullPath);
    if (stat.isDirectory()) {
      files.push(...collectFixtureXlsxFiles(fullPath));
      continue;
    }
    if (entry.toLowerCase().endsWith(".xlsx")) {
      files.push(fullPath);
    }
  }
  return files;
}

describe("xlsx2md fixture hygiene", () => {
  it("keeps local absolute path metadata out of fixture workbooks", () => {
    const fixtureRoot = path.resolve(__dirname, "./fixtures");
    const fixtureFiles = collectFixtureXlsxFiles(fixtureRoot);

    const offenders = fixtureFiles.filter((fixturePath) => {
      let workbookXml = "";
      try {
        workbookXml = execFileSync("unzip", ["-p", fixturePath, "xl/workbook.xml"], { encoding: "utf8" });
      } catch {
        return false;
      }
      return /x15ac:absPath|absPath url=/u.test(workbookXml);
    });

    expect(offenders).toEqual([]);
  });
});
