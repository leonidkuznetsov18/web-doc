import assert from "node:assert/strict";
import { describe, it } from "node:test";

import { pdfSession } from "./fixtures/pdf-session.js";

const lines = [
  "Alpha beta gamma delta",
  "Alpha beta gamma delta",
  "Iota kappa lambda mu",
] as const;

type Spacing =
  "default" | "character" | "word" | "positioned words" | "horizontal scale";

/** Raw PDF text operators exercise spacing unavailable in our PDFium builder. */
function paragraphPdf(spacing: Spacing): Uint8Array {
  const state = {
    default: "",
    character: "0.7 Tc",
    word: "4 Tw",
    "positioned words": "",
    "horizontal scale": "80 Tz",
  }[spacing];
  const content = lines
    .map((text, index) => {
      const show =
        spacing === "positioned words"
          ? `[(${text.split(" ").join(") -600 (")})] TJ`
          : `(${text}) Tj`;
      return `BT /F1 12 Tf ${state} 1 0 0 1 72 ${720 - index * 20} Tm ${show} ET`;
    })
    .join("\n");
  const objects = [
    "<< /Type /Catalog /Pages 2 0 R >>",
    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    `<< /Length ${Buffer.byteLength(content)} >>\nstream\n${content}\nendstream`,
  ];
  let pdf = "%PDF-1.7\n";
  const offsets: number[] = [];
  for (const [index, object] of objects.entries()) {
    offsets.push(Buffer.byteLength(pdf));
    pdf += `${index + 1} 0 obj\n${object}\nendobj\n`;
  }
  const xref = Buffer.byteLength(pdf);
  pdf += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  for (const offset of offsets)
    pdf += `${String(offset).padStart(10, "0")} 00000 n \n`;
  pdf += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return new Uint8Array(Buffer.from(pdf));
}

describe("native PDF paragraph spacing eligibility", () => {
  for (const spacing of [
    "character",
    "word",
    "positioned words",
    "horizontal scale",
  ] as const) {
    it(`keeps ${spacing} spacing independently editable instead of silently normalizing it`, async () => {
      const original = paragraphPdf(spacing);
      const { session, end } = await pdfSession(original);
      try {
        const elements = (await session.getElements({ pageIndex: 0 })).items;
        assert.equal(elements.length, lines.length);
        for (const element of elements) {
          assert.equal(
            element.textEditingTarget,
            undefined,
            element.text ?? element.id,
          );
          assert.equal(
            (await session.getTextParagraph(element.id)).item,
            undefined,
          );
          assert.ok(element.operations.includes("replaceText"));
        }
        assert.equal(session.state.dirty, false);
        assert.deepEqual((await session.save()).bytes, original);
      } finally {
        await end();
      }
    });
  }

  it("exposes otherwise identical normally spaced wrapped lines as one paragraph", async () => {
    const original = paragraphPdf("default");
    const { session, end } = await pdfSession(original);
    try {
      const elements = (await session.getElements({ pageIndex: 0 })).items;
      const first = elements[0];
      assert.ok(first);
      const paragraph = (await session.getTextParagraph(first.id)).item;
      assert.ok(paragraph);
      assert.equal(paragraph.text, lines.join(" "));
      assert.deepEqual(
        paragraph.memberIds,
        elements.map((element) => element.id),
      );
      assert.deepEqual((await session.save()).bytes, original);
    } finally {
      await end();
    }
  });
});
