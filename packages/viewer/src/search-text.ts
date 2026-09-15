export function normalizeSearchText(
  text: string,
  caseSensitive = false,
): string {
  const normalized = text.normalize("NFKC");
  return caseSensitive ? normalized : unicodeCaseFold(normalized);
}

export function normalizeWithMap(
  text: string,
  caseSensitive: boolean,
): { text: string; starts: number[]; ends: number[] } {
  const output: string[] = [];
  const starts: number[] = [];
  const ends: number[] = [];
  const segments = graphemeSegments(text);
  for (const segment of segments) {
    const normalized = normalizeSearchText(segment.value, caseSensitive);
    output.push(normalized);
    for (let index = 0; index < normalized.length; index += 1) {
      starts.push(segment.start);
      ends.push(segment.end);
    }
  }
  return { text: output.join(""), starts, ends };
}

export function graphemeSegments(
  text: string,
): readonly { value: string; start: number; end: number }[] {
  if (typeof Intl.Segmenter === "function") {
    const segmenter = new Intl.Segmenter(undefined, {
      granularity: "grapheme",
    });
    return [...segmenter.segment(text)].map((segment) => ({
      value: segment.segment,
      start: segment.index,
      end: segment.index + segment.segment.length,
    }));
  }
  const result: { value: string; start: number; end: number }[] = [];
  let offset = 0;
  for (const value of text) {
    result.push({ value, start: offset, end: offset + value.length });
    offset += value.length;
  }
  return result;
}

function unicodeCaseFold(text: string): string {
  return text
    .toLocaleLowerCase("und")
    .replaceAll("ß", "ss")
    .replaceAll("ς", "σ");
}
