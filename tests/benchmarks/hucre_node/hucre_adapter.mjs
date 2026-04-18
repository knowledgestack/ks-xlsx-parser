// hucre worker for the ks-xlsx-parser benchmark harness.
//
// Field names pinned against hucre 0.3.0's `_types.d.mts`:
//   sheet.rows: CellValue[][]            primitive-valued 2D array (all cells)
//   sheet.cells?: Map<string, Cell>      rich cells (only populated for cells
//                                        with style/formula/hyperlink/comment)
//   sheet.merges, dataValidations, conditionalRules, sparklines, images,
//   textBoxes, tables
//   workbook.namedRanges
//
// Protocol (matches adapters/ks_adapter.py):
//   stdout ← {"event":"ready","parser":"hucre","version":"..."}\n
//   stdin  → {"path": "...", "request_id": "..."}\n
//   stdout ← one NDJSON BenchmarkRecord per input line
//   stdin  → EOF
//   stdout ← {"event":"done"}\n

import { createInterface } from "node:readline";
import { readFileSync, statSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import { performance } from "node:perf_hooks";

const PARSER_NAME = "hucre";
const SCHEMA_VERSION = 1;
const MAX_ERR_LEN = 500;

// hucre's `exports` field does not whitelist package.json, so we read it
// directly from node_modules rather than via `require("hucre/package.json")`.
const __dirname = dirname(fileURLToPath(import.meta.url));
const HUCRE_PKG_PATH = join(__dirname, "node_modules", "hucre", "package.json");
const HUCRE_VERSION = JSON.parse(readFileSync(HUCRE_PKG_PATH, "utf8")).version;

const { readXlsx } = await import("hucre");
const HARNESS_COMMIT = process.env.HARNESS_COMMIT || "";

function write(obj) {
  process.stdout.write(JSON.stringify(obj) + "\n");
}

function emptyRecord(path, fileSize, status, errorMsg, parseTimeMs, peakMb) {
  return {
    file: path,
    file_size_bytes: fileSize,
    parser: PARSER_NAME,
    parser_version: HUCRE_VERSION,
    status,
    error: errorMsg ? String(errorMsg).slice(0, MAX_ERR_LEN) : null,
    parse_time_ms: parseTimeMs,
    peak_memory_mb: peakMb,
    sheets: null,
    cells: null,
    formulas: null,
    formula_dependencies: null,
    charts: null,
    chart_types: null,
    tables: null,
    pivots: null,
    merges: null,
    cf_rules: null,
    dv_rules: null,
    named_ranges: null,
    hyperlinks: null,
    images: null,
    comments: null,
    sparklines: null,
    chunks: null,
    token_count: null,
    schema_version: SCHEMA_VERSION,
    timestamp: new Date().toISOString(),
    harness_commit: HARNESS_COMMIT,
  };
}

// Count non-null primitive cells in a sparse row. Rows are CellValue[][] in
// hucre's shape but may be sparse objects with numeric-string keys when
// round-tripped from openXlsx.
function countRowCells(row) {
  if (row == null) return 0;
  if (Array.isArray(row)) {
    let n = 0;
    for (const v of row) if (v != null && v !== "") n += 1;
    return n;
  }
  if (typeof row === "object") {
    let n = 0;
    for (const k of Object.keys(row)) {
      if (!/^\d+$/.test(k)) continue;
      const v = row[k];
      if (v != null && v !== "") n += 1;
    }
    return n;
  }
  return 0;
}

function countFeatures(wb) {
  const sheets = wb.sheets || [];
  let cells = 0;
  let formulas = 0;
  let hyperlinks = 0;
  let comments = 0;
  let merges = 0;
  let cf = 0;
  let dv = 0;
  let images = 0;
  let sparklines = 0;
  let tables = 0;

  for (const s of sheets) {
    const rows = s.rows || [];
    for (const r of rows) cells += countRowCells(r);

    // Rich cells: Map<"row,col", Cell> — only populated with readStyles, and
    // only for cells that have style/formula/hyperlink/comment data.
    if (s.cells && typeof s.cells.forEach === "function") {
      s.cells.forEach((cell) => {
        if (cell && cell.formula != null) formulas += 1;
        if (cell && cell.hyperlink != null) hyperlinks += 1;
        if (cell && cell.comment != null) comments += 1;
      });
    }

    if (Array.isArray(s.merges)) merges += s.merges.length;
    if (Array.isArray(s.conditionalRules)) cf += s.conditionalRules.length;
    if (Array.isArray(s.dataValidations)) dv += s.dataValidations.length;
    if (Array.isArray(s.images)) images += s.images.length;
    if (Array.isArray(s.sparklines)) sparklines += s.sparklines.length;
    if (Array.isArray(s.tables)) tables += s.tables.length;
  }

  const namedRanges = Array.isArray(wb.namedRanges) ? wb.namedRanges.length : 0;

  // Charts: hucre preserves charts for round-trip via openXlsx's internal
  // state; readXlsx does not surface them on the workbook object. We cannot
  // cheaply count charts from readXlsx output, so we report null — the
  // null-vs-zero rule (_schema.py) flags this as "feature not modelled by
  // this parser at this entry point," not "zero charts found."
  const charts = null;

  // Pivots: hucre lists "No" in their feature matrix — they don't extract
  // pivot structure. We report 0 because 0 is the measured reality.
  const pivots = 0;

  return {
    sheets: sheets.length,
    cells,
    formulas,
    merges,
    cf_rules: cf,
    dv_rules: dv,
    named_ranges: namedRanges,
    hyperlinks,
    images,
    comments,
    sparklines,
    tables,
    charts,
    pivots,
  };
}

async function parseOne(path) {
  let fileSize = 0;
  try {
    fileSize = statSync(path).size;
  } catch (_) {}

  const mem0 = process.memoryUsage().rss;
  const t0 = performance.now();

  try {
    const buf = readFileSync(path);
    const wb = await readXlsx(buf, { readStyles: true });
    const t1 = performance.now();
    const mem1 = process.memoryUsage().rss;

    const counts = countFeatures(wb);
    return {
      file: path,
      file_size_bytes: fileSize,
      parser: PARSER_NAME,
      parser_version: HUCRE_VERSION,
      status: "ok",
      error: null,
      parse_time_ms: t1 - t0,
      peak_memory_mb: Math.max((mem1 - mem0) / (1024 * 1024), 0),
      sheets: counts.sheets,
      cells: counts.cells,
      formulas: counts.formulas,
      formula_dependencies: null, // hucre does not build a dep graph
      charts: counts.charts,      // null via readXlsx
      chart_types: null,          // never extracted by hucre
      tables: counts.tables,
      pivots: counts.pivots,
      merges: counts.merges,
      cf_rules: counts.cf_rules,
      dv_rules: counts.dv_rules,
      named_ranges: counts.named_ranges,
      hyperlinks: counts.hyperlinks,
      images: counts.images,
      comments: counts.comments,
      sparklines: counts.sparklines,
      chunks: null,
      token_count: null,
      schema_version: SCHEMA_VERSION,
      timestamp: new Date().toISOString(),
      harness_commit: HARNESS_COMMIT,
    };
  } catch (err) {
    const t1 = performance.now();
    const mem1 = process.memoryUsage().rss;
    return emptyRecord(path, fileSize, "error",
      `${err && err.name ? err.name : "Error"}: ${err && err.message ? err.message : String(err)}`,
      t1 - t0, Math.max((mem1 - mem0) / (1024 * 1024), 0));
  }
}

async function main() {
  write({ event: "ready", parser: PARSER_NAME, version: HUCRE_VERSION });

  const rl = createInterface({ input: process.stdin, crlfDelay: Infinity });
  for await (const line of rl) {
    const trimmed = line.trim();
    if (!trimmed) continue;
    let msg;
    try {
      msg = JSON.parse(trimmed);
    } catch (err) {
      write({ event: "error", error: `bad input line: ${err.message}` });
      continue;
    }
    const rec = await parseOne(msg.path);
    process.stdout.write(JSON.stringify(rec) + "\n");
  }

  write({ event: "done" });
}

main().catch((err) => {
  process.stderr.write(`fatal: ${err && err.stack ? err.stack : err}\n`);
  process.exit(1);
});
