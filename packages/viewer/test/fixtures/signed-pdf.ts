/*
 * Hand-written one-page PDFs with document-level features PDFium's builder
 * cannot add: a signature field (as a signed document looks to
 * FPDF_GetSignatureCount), a DocMDP certification, a tagged structure and a
 * PDF/A claim. The signature bytes themselves are empty; only the structure
 * matters.
 */

export interface HandPdfOptions {
  readonly signed?: boolean;
  readonly docMdp?: boolean;
  readonly tagged?: boolean;
  readonly pdfa?: boolean;
}

const CONTENT = "BT /F1 14 Tf 72 700 Td (Signed) Tj ET";

export function handPdf(options: HandPdfOptions = {}): Uint8Array {
  const objects: string[] = [];
  const add = (body: string): number => objects.push(body);
  // 1 catalog, filled in last; reserve its slot.
  objects.push("");
  const pages = add("<< /Type /Pages /Kids [3 0 R] /Count 1 >>");
  const page = objects.length + 1;
  const annots = options.signed ? ` /Annots [${page + 2} 0 R]` : "";
  const structParents = options.tagged ? " /StructParents 0" : "";
  add(
    `<< /Type /Page /Parent ${pages} 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 ${page + 1} 0 R >> >> /Contents ${page + 4} 0 R${annots}${structParents} >>`,
  );
  add(
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
  );
  const reference = options.docMdp
    ? " /Reference [<< /Type /SigRef /TransformMethod /DocMDP /TransformParams << /Type /TransformParams /P 1 /V /1.2 >> >>]"
    : "";
  add(
    options.signed
      ? `<< /Type /Annot /Subtype /Widget /FT /Sig /T (Signature1) /F 4 /Rect [0 0 0 0] /P ${page} 0 R /V ${page + 3} 0 R >>`
      : "<< /Type /Annot /Subtype /Link /Rect [0 0 0 0] >>",
  );
  add(
    `<< /Type /Sig /Filter /Adobe.PPKLite /SubFilter /adbe.pkcs7.detached /ByteRange [0 0 0 0] /Contents <> /Reason (Fixture) /M (D:20260101000000Z)${reference} >>`,
  );
  add(`<< /Length ${CONTENT.length} >>\nstream\n${CONTENT}\nendstream`);
  const extras: string[] = [];
  if (options.tagged) {
    const root = objects.length + 1;
    add(
      `<< /Type /StructTreeRoot /K [${root + 1} 0 R] /ParentTree << /Nums [0 [${root + 1} 0 R]] >> >>`,
    );
    add(`<< /Type /StructElem /S /P /P ${root} 0 R /Pg ${page} 0 R /K 0 >>`);
    extras.push(`/MarkInfo << /Marked true >> /StructTreeRoot ${root} 0 R`);
  }
  if (options.pdfa) {
    const xmp =
      '<?xpacket begin="" id="W5M0MpCehiHzreSzNTczkc9d"?><x:xmpmeta xmlns:x="adobe:ns:meta/"><rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#"><rdf:Description rdf:about="" xmlns:pdfaid="http://www.aiim.org/pdfa/ns/id/"><pdfaid:part>2</pdfaid:part><pdfaid:conformance>B</pdfaid:conformance></rdf:Description></rdf:RDF></x:xmpmeta><?xpacket end="w"?>';
    const metadata = add(
      `<< /Type /Metadata /Subtype /XML /Length ${xmp.length} >>\nstream\n${xmp}\nendstream`,
    );
    extras.push(`/Metadata ${metadata} 0 R`);
  }
  const acroForm = options.signed
    ? ` /AcroForm << /Fields [${page + 2} 0 R] /SigFlags 3 >>`
    : "";
  objects[0] = `<< /Type /Catalog /Pages ${pages} 0 R${acroForm}${extras.length ? ` ${extras.join(" ")}` : ""} >>`;

  let out = "%PDF-1.7\n%âãÏÓ\n";
  const offsets: number[] = [];
  objects.forEach((body, index) => {
    offsets.push(out.length);
    out += `${index + 1} 0 obj\n${body}\nendobj\n`;
  });
  const xref = out.length;
  out += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  for (const offset of offsets)
    out += `${String(offset).padStart(10, "0")} 00000 n \n`;
  out +=
    `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R ` +
    "/ID [<0123456789abcdef0123456789abcdef> <0123456789abcdef0123456789abcdef>] >>\n" +
    `startxref\n${xref}\n%%EOF\n`;
  // Every character is one byte: the file is ASCII plus the binary comment.
  return Uint8Array.from(out, (character) => character.charCodeAt(0) & 0xff);
}

/** A signed one-page PDF. */
export function signedPdf(): Uint8Array {
  return handPdf({ signed: true });
}
