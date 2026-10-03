/*
 * PDFs with text inside Form XObjects. PDFium cannot write forms, so these
 * are written out by hand: one Helvetica font, the page content, and the
 * forms as plain streams.
 */

function stream(body: string, dictionary = ""): string {
  return `<< ${dictionary} /Length ${body.length} >>\nstream\n${body}\nendstream`;
}

/** A one-page PDF from its numbered objects: 1 catalog, 2 pages, 3 page, 4 font. */
function pdf(objects: readonly string[]): Uint8Array {
  let out = "%PDF-1.4\n";
  const offsets: number[] = [];
  objects.forEach((body, index) => {
    offsets.push(out.length);
    out += `${index + 1} 0 obj\n${body}\nendobj\n`;
  });
  const xref = out.length;
  out += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  for (const offset of offsets)
    out += `${String(offset).padStart(10, "0")} 00000 n \n`;
  out += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return new Uint8Array(Buffer.from(out, "latin1"));
}

const CATALOG = "<< /Type /Catalog /Pages 2 0 R >>";
const PAGES = "<< /Type /Pages /Kids [3 0 R] /Count 1 >>";
const HELVETICA = "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>";

function page(xobjects: string, extra = ""): string {
  return `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> /XObject << ${xobjects} >> ${extra} >> /Contents 5 0 R >>`;
}

/**
 * Text on the page, and "Nested Form text" inside form Inner inside form
 * Outer drawn at (72, 650): the shape of ACTION-918's report.
 */
export function nestedFormPdf(): Uint8Array {
  return pdf([
    CATALOG,
    PAGES,
    page("/Outer 6 0 R"),
    HELVETICA,
    stream(
      "BT /F1 18 Tf 72 720 Td (Top-level control text) Tj ET\nq 1 0 0 1 72 650 cm /Outer Do Q",
    ),
    stream(
      "/Inner Do",
      "/Type /XObject /Subtype /Form /BBox [0 0 400 40] /Resources << /XObject << /Inner 7 0 R >> >>",
    ),
    stream(
      "BT /F1 24 Tf 0 10 Td (Nested Form text) Tj ET",
      "/Type /XObject /Subtype /Form /BBox [0 0 400 40] /Resources << /Font << /F1 4 0 R >> >>",
    ),
  ]);
}

/**
 * Form Outer, scaled 2× by its /Matrix and drawn at (72, 600): a red bar its
 * /BBox cuts off, "Outer own text", and form Inner drawn twice. Inner holds
 * a blue bar its own /BBox cuts off and "Shared inner".
 */
export function sharedFormPdf(): Uint8Array {
  return pdf([
    CATALOG,
    PAGES,
    page("/Outer 6 0 R"),
    HELVETICA,
    stream(
      "BT /F1 18 Tf 72 720 Td (Top-level control text) Tj ET\nq 1 0 0 1 72 600 cm /Outer Do Q",
    ),
    stream(
      "1 0 0 rg 0 0 400 8 re f\n0 g BT /F1 10 Tf 2 12 Td (Outer own text) Tj ET\nq 1 0 0 1 0 25 cm /Inner Do Q\nq 1 0 0 1 120 25 cm /Inner Do Q",
      "/Type /XObject /Subtype /Form /BBox [0 0 200 50] /Matrix [2 0 0 2 0 0] /Resources << /Font << /F1 4 0 R >> /XObject << /Inner 7 0 R >> >>",
    ),
    stream(
      "0 0 1 rg 0 0 300 4 re f\n0 g BT /F1 8 Tf 0 6 Td (Shared inner) Tj ET",
      "/Type /XObject /Subtype /Form /BBox [0 0 60 20] /Resources << /Font << /F1 4 0 R >> >>",
    ),
  ]);
}

/**
 * A form that is a transparency group drawn at half opacity, holding two
 * overlapping black squares and "Grouped text". As a group the overlap is
 * as light as the rest; drawn without the group it would come out darker.
 */
export function groupedFormPdf(): Uint8Array {
  return pdf([
    CATALOG,
    PAGES,
    page("/Group 6 0 R", "/ExtGState << /Half 7 0 R >>"),
    HELVETICA,
    stream("q /Half gs 1 0 0 1 72 600 cm /Group Do Q"),
    stream(
      "0 g 0 0 100 100 re f 50 50 100 100 re f\nBT /F1 12 Tf 0 160 Td (Grouped text) Tj ET",
      "/Type /XObject /Subtype /Form /BBox [0 0 300 200] /Group << /S /Transparency >> /Resources << /Font << /F1 4 0 R >> >>",
    ),
    "<< /Type /ExtGState /ca 0.5 /CA 0.5 >>",
  ]);
}

/**
 * Two forms drawn the way producers wrap content: one under a clip path
 * that cuts its wide bar, one at half opacity without a group. Each holds a
 * line of text.
 */
export function clippedTranslucentFormsPdf(): Uint8Array {
  return pdf([
    CATALOG,
    PAGES,
    page("/Clipped 6 0 R /Faded 7 0 R", "/ExtGState << /Half 8 0 R >>"),
    HELVETICA,
    stream(
      "q 72 600 200 60 re W n 1 0 0 1 72 600 cm /Clipped Do Q\nq /Half gs 1 0 0 1 72 400 cm /Faded Do Q",
    ),
    stream(
      "1 0 0 rg 0 0 500 10 re f\n0 g BT /F1 14 Tf 0 20 Td (Clipped text) Tj ET",
      "/Type /XObject /Subtype /Form /BBox [0 0 600 100] /Resources << /Font << /F1 4 0 R >> >>",
    ),
    stream(
      "0 0 1 rg 0 0 120 10 re f\n0 g BT /F1 14 Tf 0 20 Td (Faded text) Tj ET",
      "/Type /XObject /Subtype /Form /BBox [0 0 600 100] /Resources << /Font << /F1 4 0 R >> >>",
    ),
    "<< /Type /ExtGState /ca 0.5 /CA 0.5 >>",
  ]);
}

/**
 * A form whose text is drawn under a graphics state from the form's own
 * resources (overprint off, as many producers write before anything else).
 */
export function formGraphicsStatePdf(): Uint8Array {
  return pdf([
    CATALOG,
    PAGES,
    page("/Styled 6 0 R"),
    HELVETICA,
    stream("q 1 0 0 1 72 600 cm /Styled Do Q"),
    stream(
      "/GS0 gs 0 g BT /F1 14 Tf 0 20 Td (Styled text) Tj ET",
      "/Type /XObject /Subtype /Form /BBox [0 0 600 100] /Resources << /Font << /F1 4 0 R >> /ExtGState << /GS0 7 0 R >> >>",
    ),
    "<< /Type /ExtGState /OP false /op false /OPM 1 >>",
  ]);
}
