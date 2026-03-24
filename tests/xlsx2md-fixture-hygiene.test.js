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

  it("keeps author, timestamp, and app metadata out of fixture workbooks", () => {
    const fixtureRoot = path.resolve(__dirname, "./fixtures");
    const fixtureFiles = collectFixtureXlsxFiles(fixtureRoot);

    const offenders = fixtureFiles.filter((fixturePath) => {
      let coreXml = "";
      let appXml = "";
      try {
        coreXml = execFileSync("unzip", ["-p", fixturePath, "docProps/core.xml"], { encoding: "utf8" });
      } catch {
        coreXml = "";
      }
      try {
        appXml = execFileSync("unzip", ["-p", fixturePath, "docProps/app.xml"], { encoding: "utf8" });
      } catch {
        appXml = "";
      }
      return /<dc:creator>|<cp:lastModifiedBy>|<dcterms:created\b|<dcterms:modified\b/u.test(coreXml)
        || /<Application>|<AppVersion>/u.test(appXml);
    });

    expect(offenders).toEqual([]);
  });

  it("keeps common sensitive workbook parts out of fixture workbooks", () => {
    const fixtureRoot = path.resolve(__dirname, "./fixtures");
    const fixtureFiles = collectFixtureXlsxFiles(fixtureRoot);

    const offenders = fixtureFiles.filter((fixturePath) => {
      let zipListing = "";
      try {
        zipListing = execFileSync("unzip", ["-l", fixturePath], { encoding: "utf8" });
      } catch {
        return false;
      }
      return /xl\/comments[0-9]*\.xml|xl\/threadedComments\/|xl\/persons\/person\.xml|docProps\/custom\.xml|xl\/externalLinks\/|xl\/connections\.xml|xl\/embeddings\/|xl\/vbaProject\.bin|xl\/printerSettings\//u.test(zipListing);
    });

    expect(offenders).toEqual([]);
  });
});
