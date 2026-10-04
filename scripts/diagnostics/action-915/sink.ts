import { appendFileSync, mkdirSync, readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { execFileSync } from "node:child_process";
import { performance } from "node:perf_hooks";

const MAX_EVENTS = 8192;
const MAX_BYTES = 8 * 1024 * 1024;
const MAX_LINE_BYTES = 4096;
const fields = new Set([
  "event",
  "clock",
  "navigation",
  "format",
  "worker",
  "id",
  "operation",
  "kind",
  "timeoutMs",
  "now",
  "nodeNow",
  "jitterMs",
  "repeat",
  "status",
  "pid",
  "ppid",
  "command",
  "cpuSeconds",
  "cpuUserMs",
  "cpuSystemMs",
  "rssBytes",
  "memoryCurrentBytes",
  "memoryMaxBytes",
  "subject",
  "kit",
  "lockSha256",
  "packageSha256",
  "playwrightSha256",
  "browserType",
  "browserVersion",
  "testSha256",
  "sample",
  "failure",
  "ordinal",
  "cgroupScope",
]);

export function sanitize(
  value: unknown,
): Record<string, string | number | boolean> {
  const out: Record<string, string | number | boolean> = {};
  if (value === null || typeof value !== "object") return out;
  for (const key of fields) {
    const field: unknown = Reflect.get(value, key);
    if (typeof field === "number" && Number.isFinite(field)) out[key] = field;
    else if (typeof field === "boolean") out[key] = field;
    else if (
      typeof field === "string" &&
      field.length <= 128 &&
      /^[a-zA-Z0-9_.:/-]+$/.test(field)
    )
      out[key] = field;
  }
  return out;
}

export function createSink(output: string) {
  mkdirSync(dirname(output), { recursive: true });
  let events = 0;
  let bytes = 0;
  let capped = false;
  const clock = `node-${process.pid}`;
  return (value: unknown): void => {
    if (capped) return;
    const line =
      JSON.stringify({
        ...sanitize(value),
        sinkClock: clock,
        sinkNow: performance.now(),
        ordinal: ++events,
      }) + "\n";
    const length = Buffer.byteLength(line);
    if (
      events > MAX_EVENTS ||
      bytes + length > MAX_BYTES ||
      length > MAX_LINE_BYTES
    ) {
      appendFileSync(
        output,
        JSON.stringify({
          event: "capture-cap",
          sinkClock: clock,
          sinkNow: performance.now(),
        }) + "\n",
      );
      capped = true;
      return;
    }
    appendFileSync(output, line);
    bytes += length;
  };
}

function memory(path: string): number | "max" | "unavailable" {
  try {
    const value = readFileSync(path, "utf8").trim();
    if (value === "max") return value;
    const parsed = Number(value);
    return Number.isSafeInteger(parsed) && parsed >= 0 ? parsed : "unavailable";
  } catch {
    return "unavailable";
  }
}

/** Cumulative OS process CPU/RSS only; this is not a WebKit JS-heap measurement. */
export function sampleResources(record: (value: unknown) => void): void {
  const cpu = process.cpuUsage();
  let group: string | undefined;
  try {
    const match = /^0::(\/[^\n]*)$/m.exec(
      readFileSync("/proc/self/cgroup", "utf8"),
    );
    if (match && !match[1].split("/").includes("..")) group = match[1];
  } catch {
    record({
      event: "cgroup-unavailable",
      status: "self-membership-unavailable",
    });
  }
  const directory = group ? resolve("/sys/fs/cgroup", `.${group}`) : undefined;
  record({
    event: "node-resource",
    cpuUserMs: cpu.user / 1000,
    cpuSystemMs: cpu.system / 1000,
    rssBytes: process.memoryUsage().rss,
    cgroupScope:
      group === "/"
        ? "root-shared"
        : group
          ? "self-membership-path"
          : "unverified",
    memoryCurrentBytes: directory
      ? memory(`${directory}/memory.current`)
      : "unavailable",
    memoryMaxBytes: directory
      ? memory(`${directory}/memory.max`)
      : "unavailable",
  });
  if (process.platform !== "linux") {
    record({ event: "os-resource-unavailable", status: "non-linux" });
    return;
  }
  try {
    // No command arguments, environment, paths or process output is retained.
    const rows = execFileSync(
      "ps",
      ["-eo", "pid,ppid,comm,time,rss", "--no-headers"],
      { encoding: "utf8", timeout: 1000, maxBuffer: 128 * 1024 },
    )
      .trim()
      .split("\n");
    let sampled = 0;
    for (const row of rows) {
      const match = /^\s*(\d+)\s+(\d+)\s+(\S+)\s+([\d:-]+)\s+(\d+)\s*$/.exec(
        row,
      );
      if (
        !match ||
        !/^(node|WebKit[A-Za-z]*|MiniBrowser|Playwright.*)$/.test(match[3])
      )
        continue;
      if (sampled++ === 32) break;
      const time = match[4].split(/[-:]/).map(Number);
      if (time.length > 4) continue;
      const [seconds = 0, minutes = 0, hours = 0, days = 0] = time.reverse();
      const cpuSeconds = seconds + minutes * 60 + hours * 3600 + days * 86400;
      record({
        event: "os-process",
        pid: Number(match[1]),
        ppid: Number(match[2]),
        command: match[3],
        cpuSeconds,
        rssBytes: Number(match[5]) * 1024,
      });
    }
  } catch {
    record({ event: "os-resource-unavailable", status: "sample-failed" });
  }
}
