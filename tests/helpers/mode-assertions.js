import { expect } from "vitest";

export function expectModeResults(runMode, expectedByMode) {
  expect(runMode("plain")).toEqual(expectedByMode.plain);
  expect(runMode("github")).toEqual(expectedByMode.github);
}

