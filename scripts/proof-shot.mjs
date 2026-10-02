import { readFileSync } from "node:fs";
import { chromium } from "@playwright/test";

/*
 * Renders text from stdin (or a file) as a PNG, for test-run proofs attached
 * to tickets:  node scripts/proof-shot.mjs --title "T23" --out proof.png < run.txt
 */

const args = process.argv.slice(2);
const option = (name, fallback) => {
  const index = args.indexOf(name);
  return index >= 0 ? args[index + 1] : fallback;
};
const title = option("--title", "Test run");
const out = option("--out", "proof.png");
const input = option("--in");
const text = input ? readFileSync(input, "utf8") : readFileSync(0, "utf8");

const escape = (value) =>
  value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
const html = `<!doctype html><meta charset="utf-8"><body style="margin:0;background:#0f172a;color:#e2e8f0;font:13px/1.45 ui-monospace,Menlo,monospace">
<div style="padding:14px 18px;background:#1e293b;font:600 15px system-ui,sans-serif;color:#f8fafc">${escape(title)}<span style="float:right;font-weight:400;color:#94a3b8">${new Date().toISOString().slice(0, 16).replace("T", " ")} UTC</span></div>
<pre style="margin:0;padding:16px 18px;white-space:pre-wrap">${escape(text.trimEnd())}</pre></body>`;

const browser = await chromium.launch();
const page = await browser.newPage({ viewport: { width: 960, height: 400 } });
await page.setContent(html);
await page.screenshot({ path: out, fullPage: true });
await browser.close();
console.log(out);
