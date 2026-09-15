import { expect, test } from "@playwright/test";
import { citationBoundaryScenario } from "./citation-scenario.js";

for (const mode of ["worker", "main", "unavailable"] as const) {
  test(`citation highlights stop before the next DOCX run (${mode})`, async ({
    page,
  }) => {
    await page.goto("/");
    const result = await page.evaluate(citationBoundaryScenario, mode);
    expect(result.strategy).toBe("fuzzy");
    expect(result.highlightCount).toBe(1);
  });
}
