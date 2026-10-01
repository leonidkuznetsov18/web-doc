/*
 * A hand-written one-page PDF with a signature field, as a signed document
 * looks to PDFium: an AcroForm whose field has /FT /Sig and a /V signature
 * dictionary. The signature bytes themselves are empty; only the structure
 * matters for FPDF_GetSignatureCount.
 */

const CONTENT = "BT /F1 14 Tf 72 700 Td (Signed) Tj ET";

const OBJECTS = [
  "<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [5 0 R] /SigFlags 3 >> >>",
  "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
  "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 7 0 R >> >> /Contents 4 0 R /Annots [5 0 R] >>",
  `<< /Length ${CONTENT.length} >>\nstream\n${CONTENT}\nendstream`,
  "<< /Type /Annot /Subtype /Widget /FT /Sig /T (Signature1) /F 4 /Rect [0 0 0 0] /P 3 0 R /V 6 0 R >>",
  "<< /Type /Sig /Filter /Adobe.PPKLite /SubFilter /adbe.pkcs7.detached /ByteRange [0 0 0 0] /Contents <> /Reason (Fixture) /M (D:20260101000000Z) >>",
  "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
];

export function signedPdf(): Uint8Array {
  let out = "%PDF-1.7\n%âãÏÓ\n";
  const offsets: number[] = [];
  OBJECTS.forEach((body, index) => {
    offsets.push(out.length);
    out += `${index + 1} 0 obj\n${body}\nendobj\n`;
  });
  const xref = out.length;
  out += `xref\n0 ${OBJECTS.length + 1}\n0000000000 65535 f \n`;
  for (const offset of offsets)
    out += `${String(offset).padStart(10, "0")} 00000 n \n`;
  out +=
    `trailer\n<< /Size ${OBJECTS.length + 1} /Root 1 0 R ` +
    "/ID [<0123456789abcdef0123456789abcdef> <0123456789abcdef0123456789abcdef>] >>\n" +
    `startxref\n${xref}\n%%EOF\n`;
  // Every character is one byte: the file is ASCII plus the binary comment.
  return Uint8Array.from(out, (character) => character.charCodeAt(0) & 0xff);
}
