import { mkdir, writeFile } from "node:fs/promises";
import { resolve } from "node:path";
import { performance } from "node:perf_hooks";

import {
  defaultResourceLimits,
  detectFormat,
  enforceContainerLimits,
  parseDelimitedBytes,
  sanitizeSvg,
} from "../packages/viewer/dist/index.js";
import { compactPdf } from "../packages/viewer/dist/edit/pdf/engine/compact.js";
import {
  localRecordOf,
  parseZip,
} from "../packages/viewer/dist/edit/ooxml/zip.js";

const root = resolve(import.meta.dirname, "..");
const iterations = Number(process.env.FUZZ_ITERATIONS ?? 2_000);
const maxCaseMs = Number(process.env.FUZZ_CASE_MS ?? 100);
const seeds = [
  bytes("%PDF-1.7\n1 0 obj<<>>endobj"),
  bytes(
    "%PDF-1.7\n1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents 4 0 R >>\nendobj\n4 0 obj\n<< /Length 5 0 R >>\nstream\nBT (x) Tj ET\nendstream\nendobj\n5 0 obj\n12\nendobj\n6 0 obj\n<< /Orphan (endobj stream) >>\nendobj\nxref\n0 7\n0000000000 65535 f \ntrailer\n<< /Size 7 /Root 1 0 R >>\nstartxref\n9\n%%EOF\n",
  ),
  Uint8Array.of(0x50, 0x4b, 0x03, 0x04, 0, 0, 0, 0),
  tinyZip(),
  Uint8Array.of(0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1),
  bytes('<svg xmlns="http://www.w3.org/2000/svg"><script>x</script></svg>'),
  bytes('a,b\n"quoted\nfield",c'),
  Uint8Array.of(0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a),
];

let randomState = 0x5eed1234;
let slowestMs = 0;
let failures = 0;
const startedAt = performance.now();

for (let iteration = 0; iteration < iterations; iteration += 1) {
  const sample = mutate(seeds[iteration % seeds.length]);
  const caseStarted = performance.now();
  exercise(sample);
  const elapsed = performance.now() - caseStarted;
  slowestMs = Math.max(slowestMs, elapsed);
  if (elapsed > maxCaseMs) {
    failures += 1;
    throw new Error(
      `Mutation ${iteration} exceeded ${maxCaseMs} ms (${elapsed.toFixed(2)} ms)`,
    );
  }
}

const report = {
  schemaVersion: 1,
  seed: "0x5eed1234",
  iterations,
  failures,
  elapsedMs: Number((performance.now() - startedAt).toFixed(2)),
  slowestCaseMs: Number(slowestMs.toFixed(2)),
  targets: [
    "detection/ZIP/CFB",
    "SVG sanitizer",
    "CSV/TSV parser",
    "PDF full-save compaction",
    "OOXML ZIP reader",
  ],
};
await mkdir(resolve(root, "artifacts"), { recursive: true });
await writeFile(
  resolve(root, "artifacts/fuzz-js.json"),
  `${JSON.stringify(report, null, 2)}\n`,
);
console.log("JavaScript mutation fuzz passed", report);

function exercise(sample) {
  try {
    const detection = detectFormat(sample, { fileName: "mutation.csv" });
    try {
      enforceContainerLimits(sample, detection.format, {
        ...defaultResourceLimits,
        maxInputBytes: 2 * 1024 * 1024,
      });
    } catch {}
  } catch {}
  const text = new TextDecoder("utf-8", { fatal: false }).decode(sample);
  try {
    sanitizeSvg(text);
  } catch {}
  for (const format of ["csv", "tsv"])
    try {
      parseDelimitedBytes(sample, format, 10_000);
    } catch {}
  // The compaction pass of a full PDF save must refuse or finish, never hang.
  try {
    compactPdf(sample);
  } catch {}
  // The OOXML ZIP reader must refuse or finish on any bytes.
  try {
    const archive = parseZip(sample, defaultResourceLimits);
    for (const entry of archive.entries)
      try {
        localRecordOf(archive, entry);
      } catch {}
  } catch {}
}

/** A stored one-entry archive, the seed for the ZIP reader. */
function tinyZip() {
  const name = bytes("a.xml");
  const data = bytes("<a/>");
  const local = new Uint8Array(30 + name.length);
  const view = new DataView(local.buffer);
  view.setUint32(0, 0x04034b50, true);
  view.setUint16(4, 20, true);
  view.setUint32(14, 0x1a8e3f3c, true);
  view.setUint32(18, data.length, true);
  view.setUint32(22, data.length, true);
  view.setUint16(26, name.length, true);
  local.set(name, 30);
  const central = new Uint8Array(46 + name.length);
  const centralView = new DataView(central.buffer);
  centralView.setUint32(0, 0x02014b50, true);
  centralView.setUint16(6, 20, true);
  centralView.setUint32(16, 0x1a8e3f3c, true);
  centralView.setUint32(20, data.length, true);
  centralView.setUint32(24, data.length, true);
  centralView.setUint16(28, name.length, true);
  central.set(name, 46);
  const eocd = new Uint8Array(22);
  const eocdView = new DataView(eocd.buffer);
  eocdView.setUint32(0, 0x06054b50, true);
  eocdView.setUint16(8, 1, true);
  eocdView.setUint16(10, 1, true);
  eocdView.setUint32(12, central.length, true);
  eocdView.setUint32(16, local.length + data.length, true);
  const out = new Uint8Array(local.length + data.length + central.length + 22);
  out.set(local, 0);
  out.set(data, local.length);
  out.set(central, local.length + data.length);
  out.set(eocd, local.length + data.length + central.length);
  return out;
}

function mutate(seed) {
  let output = seed.slice();
  const operations = 1 + (random() % 8);
  for (let operation = 0; operation < operations; operation += 1) {
    const choice = random() % 4;
    if (choice === 0 && output.length > 0) {
      output[random() % output.length] ^= random() & 0xff;
    } else if (choice === 1 && output.length > 0) {
      output = output.slice(0, random() % output.length);
    } else if (choice === 2 && output.length < 64 * 1024) {
      const extra = new Uint8Array(1 + (random() % 128));
      for (let index = 0; index < extra.length; index += 1)
        extra[index] = random() & 0xff;
      const joined = new Uint8Array(output.length + extra.length);
      joined.set(output);
      joined.set(extra, output.length);
      output = joined;
    } else if (output.length > 0 && output.length < 32 * 1024) {
      const joined = new Uint8Array(output.length * 2);
      joined.set(output);
      joined.set(output, output.length);
      output = joined;
    }
  }
  return output;
}

function random() {
  randomState = (Math.imul(randomState, 1_664_525) + 1_013_904_223) >>> 0;
  return randomState;
}

function bytes(value) {
  return new TextEncoder().encode(value);
}
