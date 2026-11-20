import { test, after } from "node:test";
import assert from "node:assert/strict";
import { readFileSync, writeFileSync, openSync, closeSync, fstatSync, mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { createRequire } from "node:module";
// Import by package name so tests go through the package.json "exports" map
import init, { Workbook, LazyWorkbook } from "@kuramajs/excel";
import { writeXlsx } from "@kuramajs/excel/writer";

await init();
const bytes = readFileSync(new URL("./fixtures/basic.xlsx", import.meta.url));

test("lists worksheets with visibility", () => {
  const wb = Workbook.open(bytes);
  assert.deepEqual(wb.sheets(), [
    { name: "Data", visibility: "Visible" },
    { name: "Secret", visibility: "Hidden" },
    { name: "People", visibility: "Visible" },
    { name: "Offset", visibility: "Visible" },
  ]);
});

test("returns typed cells and UTC dates", () => {
  const wb = Workbook.open(bytes);
  const sheet = wb.sheet("Data");
  assert.equal(sheet.height, 2);
  assert.equal(sheet.width, 5);
  const [first, second] = sheet.rows();
  assert.deepEqual(first, ["name", 42.5, true, new Date("2024-01-02T13:14:15Z"), { error: "#DIV/0!" }]);
  assert.deepEqual(second, [null, 7, null, null, null]);
  assert.deepEqual(sheet.row(1), second);
});

test("accepts ArrayBuffer", () => {
  const buf = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
  const wb = Workbook.open(buf);
  assert.equal(wb.sheets().length, 4);
});

test("throws errors with codes", () => {
  const code = (code) => (err) => err instanceof Error && err.code === code;
  assert.throws(() => Workbook.open(new Uint8Array([1, 2, 3])), code("INVALID_FILE"));
  assert.throws(() => Workbook.open("nope"), code("INVALID_INPUT"));
  const wb = Workbook.open(bytes);
  assert.throws(() => wb.sheet("Missing"), code("SHEET_NOT_FOUND"));
  assert.throws(() => wb.forEachRow("Missing", () => {}), code("SHEET_NOT_FOUND"));
  const sheet = wb.sheet("Data");
  assert.throws(() => sheet.row(99), code("OUT_OF_BOUNDS"));
  assert.throws(() => sheet.toObjects({ header: 99 }), code("OUT_OF_BOUNDS"));
});

test("toObjects keys rows by header, dedupes and fills blanks", () => {
  const wb = Workbook.open(bytes);
  assert.deepEqual(wb.sheet("People").toObjects(), [
    { id: 1, name: "Ann", __EMPTY_2: null, name_1: null },
    { id: null, name: null, __EMPTY_2: null, name_1: null },
    { id: 3, name: null, __EMPTY_2: null, name_1: null },
  ]);
  assert.deepEqual(wb.sheet("Data").toObjects({ header: 1 }), []);
});

test("forEachRow streams same rows as rows(), including gaps", () => {
  const wb = Workbook.open(bytes);
  for (const name of ["Data", "People", "Offset"]) {
    const streamed = [];
    wb.forEachRow(name, (row, i) => { streamed[i] = row; });
    assert.deepEqual(streamed, wb.sheet(name).rows(), name);
    assert.deepEqual([...wb.iterRows(name)], wb.sheet(name).rows(), name);
  }
});

test("iterRows supports early break and throws eagerly", () => {
  const wb = Workbook.open(bytes);
  let count = 0;
  for (const row of wb.iterRows("People")) {
    count++;
    assert.deepEqual(row, ["id", "name", null, "name"]);
    break;
  }
  assert.equal(count, 1);
  const it = wb.iterRows("People");
  it.next();
  assert.deepEqual(it.return(), { done: true, value: undefined });
  assert.deepEqual(it.next(), { done: true, value: undefined });
  assert.throws(() => wb.iterRows("Missing"), (err) => err.code === "SHEET_NOT_FOUND");
});

test("streamed rows stay rectangular when the <dimension> tag is wrong", () => {
  const wb = Workbook.open(readFileSync(new URL("./fixtures/bad_dimension.xlsx", import.meta.url)));
  for (const name of ["Data", "People"]) {
    const expected = wb.sheet(name).rows();
    const streamed = [];
    wb.forEachRow(name, (row, i) => { streamed[i] = row; });
    assert.deepEqual(streamed, expected, name);
    assert.deepEqual([...wb.iterRows(name)], expected, name);
  }
});

test("iterRows hands the workbook back, so repeated and overlapping use works", () => {
  const wb = Workbook.open(bytes);
  const expected = wb.sheet("People").rows();

  // sequential: finished, early-broken and freed iterators each return the workbook
  for (let i = 0; i < 3; i++) assert.deepEqual([...wb.iterRows("People")], expected);
  for (const _ of wb.iterRows("Data")) break;
  const started = wb.iterRows("Data");
  started.next();
  started.return();
  assert.deepEqual(wb.sheet("Offset").rows()[2], [null, null, "C3", null]);

  // overlapping: other calls work while an iterator is open
  const it = wb.iterRows("People");
  assert.deepEqual(it.next().value, expected[0]);
  assert.deepEqual(wb.sheets().map((s) => s.name), ["Data", "Secret", "People", "Offset"]);
  assert.deepEqual(wb.sheet("People").rows(), expected);
  assert.deepEqual([...wb.iterRows("Data")], wb.sheet("Data").rows());
  assert.deepEqual([...it], expected.slice(1));

  // and the workbook is still usable afterwards
  assert.deepEqual([...wb.iterRows("People")], expected);
  assert.throws(() => wb.iterRows("Missing"), (err) => err.code === "SHEET_NOT_FOUND");
  assert.deepEqual([...wb.iterRows("People")], expected);
});

test("indexes rows and columns from A1", () => {
  const wb = Workbook.open(bytes);
  const sheet = wb.sheet("Offset");
  assert.deepEqual(sheet.rows(), [
    [null, null, null, null],
    [null, null, null, null],
    [null, null, "C3", null],
    [null, null, null, "D4"],
  ]);
  assert.equal(sheet.row(2)[2], "C3");
});

test("forEachRow stops on false and propagates throws", () => {
  const wb = Workbook.open(bytes);
  const seen = [];
  wb.forEachRow("People", (_row, i) => { seen.push(i); return false; });
  assert.deepEqual(seen, [0]);
  assert.throws(() => wb.forEachRow("People", () => { throw new Error("boom"); }), /boom/);
});

test("supports explicit and Symbol.dispose cleanup", () => {
  const wb = Workbook.open(bytes);
  assert.equal(typeof wb[Symbol.dispose], "function");
  wb.free();
});

const fixture = (name) => readFileSync(new URL(`./fixtures/${name}`, import.meta.url));
const d = (iso) => new Date(iso);

for (const file of ["date.xls", "date_1904.xls", "date.xlsb"]) {
  test(`reads dates and durations from ${file}`, () => {
    const wb = Workbook.open(fixture(file));
    const [{ name }] = wb.sheets();
    const rows = wb.sheet(name).rows();
    assert.deepEqual(rows, [
      [d("2021-01-01T00:00:00Z"), 15],
      [d("2021-01-02T00:00:00Z"), 16],
      [918610, 17],
    ]);
    const streamed = [];
    wb.forEachRow(name, (row, i) => { streamed[i] = row; });
    assert.deepEqual(streamed, rows);
    assert.deepEqual([...wb.iterRows(name)], rows);
  });
}

test("reads ISO dates and durations from ods", () => {
  const wb = Workbook.open(fixture("date.ods"));
  const [{ name }] = wb.sheets();
  assert.deepEqual(wb.sheet(name).rows(), [
    [d("2021-01-01T00:00:00Z"), 15],
    [d("2021-01-01T10:10:10Z"), 16],
    [36610, 17],
    [36610.123456, 18],
  ]);
});

for (const file of ["pass_protected.xlsx", "pass_protected.xlsb", "pass_protected.ods"]) {
  test(`reports PASSWORD_PROTECTED for ${file}`, () => {
    assert.throws(() => Workbook.open(fixture(file)), (err) => err.code === "PASSWORD_PROTECTED");
  });
}

test("sheet onProgress gives same rows and reports progress", () => {
  const wb = Workbook.open(bytes);
  for (const name of ["Data", "People", "Offset"]) {
    const calls = [];
    const sheet = wb.sheet(name, { onProgress: (read, total) => calls.push([read, total]) });
    assert.deepEqual(sheet.rows(), wb.sheet(name).rows(), name);
    assert.ok(calls.length >= 2, name);
    const [lastRead, lastTotal] = calls.at(-1);
    assert.equal(lastRead, lastTotal, name);
  }
  assert.throws(() => wb.sheet("Missing", { onProgress() {} }), (err) => err.code === "SHEET_NOT_FOUND");
});

test("rowsSlice behaves like Array.slice, clamped", () => {
  const wb = Workbook.open(bytes);
  const sheet = wb.sheet("People");
  const rows = sheet.rows();
  assert.deepEqual(sheet.rowsSlice(1, 3), rows.slice(1, 3));
  assert.deepEqual(sheet.rowsSlice(2, 99), rows.slice(2));
  assert.deepEqual(sheet.rowsSlice(3, 1), []);
  assert.deepEqual(sheet.rowsSlice(50, 60), []);
});

test("reads csv as one sheet of strings", () => {
  const csv = Buffer.from('name,age\nAnn,30\n"Smith, J","two\nlines"\r\nBob\n\n"say ""hi""",01234\n');
  const wb = Workbook.open(csv);
  assert.deepEqual(wb.sheets(), [{ name: "Sheet1", visibility: "Visible" }]);
  const rows = [
    ["name", "age"],
    ["Ann", "30"],
    ["Smith, J", "two\nlines"],
    ["Bob", null],
    ['say "hi"', "01234"],
  ];
  assert.deepEqual(wb.sheet("Sheet1").rows(), rows);
  assert.deepEqual([...wb.iterRows("Sheet1")], rows);
  const streamed = [];
  wb.forEachRow("Sheet1", (row) => void streamed.push(row));
  assert.deepEqual(streamed, rows);
  assert.deepEqual(wb.sheet("Sheet1").toObjects()[0], { name: "Ann", age: "30" });
  assert.throws(() => wb.sheet("Data"), (e) => e.code === "SHEET_NOT_FOUND");
});

test("csv detects delimiter, BOM and windows-1252", () => {
  const rows = (b) => Workbook.open(b).sheet("Sheet1").rows();
  assert.deepEqual(rows(Buffer.from("﻿a;b\n1,5;2\n")), [["a", "b"], ["1,5", "2"]]);
  assert.deepEqual(rows(Buffer.from("a\tb\n1\t2")), [["a", "b"], ["1", "2"]]);
  assert.deepEqual(rows(Buffer.from([0x63, 0x61, 0x66, 0xe9])), [["café"]]);
});

test("rejects input that is neither a workbook nor csv", () => {
  const invalid = (e) => e.code === "INVALID_FILE";
  assert.throws(() => Workbook.open(new Uint8Array()), invalid);
  assert.throws(() => Workbook.open(Buffer.from("<html><body>x</body></html>")), invalid);
});

test("browser init() explains an HTML page served instead of the wasm", async () => {
  const { default: webInit } = await import("../js/web.mjs");
  const html = new Response("<!DOCTYPE html><html></html>", { headers: { "Content-Type": "text/html" } });
  const warn = console.warn;
  console.warn = () => {}; // wasm-bindgen warns about the wrong MIME type before failing
  try {
    await assert.rejects(webInit({ module_or_path: html }), (e) => {
      assert.match(e.message, /returned an HTML page.*excel\/troubleshooting/);
      assert.ok(e.cause instanceof WebAssembly.CompileError);
      return true;
    });
  } finally {
    console.warn = warn;
  }
});

test("CommonJS require exposes the same API", () => {
  const cjs = createRequire(import.meta.url)("@kuramajs/excel");
  const wb = cjs.Workbook.open(bytes);
  assert.equal(wb.sheets().length, 4);
  assert.deepEqual(Object.keys(cjs).sort(), ["LazyWorkbook", "RowIter", "Sheet", "Workbook"]);
  const lazy = cjs.LazyWorkbook.open(new URL("./fixtures/basic.xlsx", import.meta.url).pathname);
  assert.deepEqual(lazy.sheet("Data").rows(), wb.sheet("Data").rows());
  lazy.free();
});

// Everything a caller can read from a workbook, for comparing lazy against in-memory opens.
function dump(wb) {
  return wb.sheets().map(({ name }) => {
    const sheet = wb.sheet(name);
    const streamed = [];
    wb.forEachRow(name, (row) => void streamed.push(row));
    return { name, rows: sheet.rows(), iter: [...wb.iterRows(name)], streamed };
  });
}

const fixturePath = (name) => new URL(`./fixtures/${name}`, import.meta.url).pathname;
const csvText = "name,age\nAnn,30\nBob,41\n";
const tmp = mkdtempSync(join(tmpdir(), "kurama-lazy-"));
after(() => rmSync(tmp, { recursive: true, force: true }));
writeFileSync(join(tmp, "small.csv"), csvText);
// > 1 MB block size, so reads cross blocks and the block cache gets used
const bigRows = Array.from({ length: 150_000 }, (_, i) => ({ id: i, name: `row ${i} ${"x".repeat(20)}`, v: i / 7 }));
writeFileSync(join(tmp, "big.xlsx"), writeXlsx({ Big: bigRows, Second: [{ a: 1 }] }));

for (const path of [
  ...["basic.xlsx", "bad_dimension.xlsx", "date.ods", "date.xls", "date_1904.xls", "date.xlsb"].map(fixturePath),
  join(tmp, "small.csv"),
  join(tmp, "big.xlsx"),
]) {
  test(`LazyWorkbook reads ${path.split("/").pop()} like Workbook.open`, () => {
    const lazy = LazyWorkbook.open(path);
    try {
      assert.deepEqual(dump(lazy), dump(Workbook.open(readFileSync(path))));
    } finally {
      lazy[Symbol.dispose]();
    }
  });
}

test("LazyWorkbook keeps password errors and validates the source", () => {
  assert.throws(() => LazyWorkbook.open(fixturePath("pass_protected.xlsx")), (e) => e.code === "PASSWORD_PROTECTED");
  assert.throws(() => Workbook.openSource({}), (e) => e.code === "INVALID_INPUT");
  assert.throws(() => Workbook.openSource({ size: -1, readAt() {} }), (e) => e.code === "INVALID_INPUT");
});

test("openSource reports a short or failing read as READ_FAILED", () => {
  const big = readFileSync(join(tmp, "big.xlsx"));
  const short = { size: big.length + 10, readAt: (o, n) => big.subarray(o, o + n) };
  assert.throws(() => Workbook.openSource(short), (e) => e.code === "READ_FAILED");
  const throws = { size: big.length, readAt() { throw new Error("disk gone"); } };
  assert.throws(() => Workbook.openSource(throws), (e) => e.code === "READ_FAILED" && /disk gone/.test(e.message));
});

test("openSource only reads the blocks it needs", () => {
  const big = readFileSync(join(tmp, "big.xlsx"));
  let read = 0;
  assert.ok(big.length > 2 << 20, "fixture must span several 1 MB blocks");
  const wb = Workbook.openSource({ size: big.length, readAt: (o, n) => ((read += n), big.subarray(o, o + n)) });
  assert.deepEqual(wb.sheet("Second").rows(), [["a"], [1]]);
  assert.ok(read < big.length, `read ${read} of ${big.length} bytes`);
});

test("LazyWorkbook closes its file when the workbook is garbage collected without free()", async () => {
  const { setFlagsFromString } = await import("node:v8");
  const { runInNewContext } = await import("node:vm");
  setFlagsFromString("--expose-gc");
  const gc = runInNewContext("gc");
  const { readdirSync } = await import("node:fs");
  const openFds = () => readdirSync("/dev/fd").length;
  const before = openFds();
  (() => {
    const wb = LazyWorkbook.open(fixturePath("basic.xlsx"));
    assert.equal(wb.sheets().length, 4);
  })();
  assert.equal(openFds(), before + 1);
  // Two hops: wasm-bindgen frees the Rust workbook, which releases readAt, which closes the file.
  for (let i = 0; i < 50 && openFds() > before; i++) {
    gc();
    await new Promise((r) => setTimeout(r, 10));
  }
  assert.equal(openFds(), before);
});

test("LazyWorkbook leaves a caller's fd open", () => {
  const fd = openSync(fixturePath("basic.xlsx"), "r");
  try {
    LazyWorkbook.open(fd).free();
    assert.ok(fstatSync(fd).size > 0);
  } finally {
    closeSync(fd);
  }
});

test("browser LazyWorkbook needs a worker", async () => {
  const web = await import("../js/web.mjs");
  assert.throws(() => web.LazyWorkbook.open(new Blob([bytes])), (e) => e.code === "NOT_IN_WORKER");
});

test("browser LazyWorkbook reads a Blob through FileReaderSync", async () => {
  const web = await import("../js/web.mjs");
  await web.default({ module_or_path: readFileSync(new URL("../pkg/web/web_index_bg.wasm", import.meta.url)) });
  // Node has no sync Blob read, so record each slice's bytes and fake the worker-only FileReaderSync.
  const blobBytes = new WeakMap();
  globalThis.FileReaderSync = class {
    readAsArrayBuffer(blob) {
      return blobBytes.get(blob);
    }
  };
  const big = readFileSync(join(tmp, "big.xlsx"));
  const blob = new Blob([big]);
  const origSlice = Blob.prototype.slice;
  Blob.prototype.slice = function (a, b) {
    const part = origSlice.call(this, a, b);
    blobBytes.set(part, big.buffer.slice(big.byteOffset + a, big.byteOffset + Math.min(b, big.length)));
    return part;
  };
  try {
    const wb = web.LazyWorkbook.open(blob);
    assert.deepEqual(wb.sheet("Second").rows(), [["a"], [1]]);
  } finally {
    Blob.prototype.slice = origSlice;
    delete globalThis.FileReaderSync;
  }
});
