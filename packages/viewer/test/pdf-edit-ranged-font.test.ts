import assert from "node:assert/strict";
import { it } from "node:test";

import { extractPageText } from "./fixtures/pdf-builder.js";
import { pdfSession } from "./fixtures/pdf-session.js";

it("removes an unsupported extracted character without fetching a font for removed text", async () => {
  const { session, end } = await pdfSession(controlCharacterPdf());
  try {
    const original = (await session.getElement("p0:o0")).item;
    assert.equal(original?.text, "Hello\u0002world");
    const receipt = await session.replaceText({
      target: "p0:o0",
      text: " ",
      range: {
        start: { elementId: "p0:o0", offset: 5 },
        end: { elementId: "p0:o0", offset: 6 },
      },
    });
    assert.equal((await session.getElement("p0:o0")).item?.text, "Hello world");
    assert.deepEqual(receipt.warnings, []);
    assert.equal(
      await extractPageText((await session.save()).bytes, 0),
      "Hello world",
    );
    await session.undo();
    assert.equal(
      (await session.getElement("p0:o0")).item?.text,
      original?.text,
    );
  } finally {
    await end();
  }
});

/** A real ToUnicode map makes PDFium expose a printed hyphen as U+0002. */
function controlCharacterPdf(): Uint8Array {
  const content = "BT /F1 12 Tf 1 0 0 1 72 700 Tm (Hello-world) Tj ET";
  const cmap = `/CIDInit /ProcSet findresource begin
12 dict begin
begincmap
/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def
/CMapName /HyphenControl def
/CMapType 2 def
1 begincodespacerange
<00> <FF>
endcodespacerange
1 beginbfchar
<2D> <0002>
endbfchar
endcmap
CMapName currentdict /CMap defineresource pop
end end`;
  const stream = (text: string): string =>
    `<< /Length ${Buffer.byteLength(text)} >>\nstream\n${text}\nendstream`;
  const objects = [
    "<< /Type /Catalog /Pages 2 0 R >>",
    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /ToUnicode 6 0 R >>",
    stream(content),
    stream(cmap),
  ];
  let text = "%PDF-1.7\n";
  const offsets = [0];
  objects.forEach((object, index) => {
    offsets.push(Buffer.byteLength(text));
    text += `${index + 1} 0 obj\n${object}\nendobj\n`;
  });
  const xref = Buffer.byteLength(text);
  text += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  for (const offset of offsets.slice(1))
    text += `${String(offset).padStart(10, "0")} 00000 n \n`;
  text += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return new Uint8Array(Buffer.from(text));
}
