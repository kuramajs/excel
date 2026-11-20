// Prepend `/// <reference lib>` directives to the generated typings, so consumers don't need
// matching `lib` settings in their tsconfig.
const fs = require("fs");

const refs = {
  // esnext.disposable: Symbol.dispose on Workbook and Sheet
  "pkg/web/web_index.d.ts": ["esnext.disposable", "dom"], // dom: init() takes fetch/WebAssembly types
  "pkg/nodejs/nodejs_index.d.ts": ["esnext.disposable"],
  "pkg/writer/web/writer.d.ts": ["esnext.disposable", "dom"],
  "pkg/writer/nodejs/writer.d.ts": ["esnext.disposable"],
};

for (const [file, libs] of Object.entries(refs)) {
  const header = libs.map((lib) => `/// <reference lib="${lib}" />\n`).join("");
  fs.writeFileSync(file, header + fs.readFileSync(file, "utf8"));
}

// LazyWorkbook lives in js/web.mjs and js/node.mjs (not wasm-bindgen), so declare it here.
fs.appendFileSync(
  "pkg/web/web_index.d.ts",
  `
/**
 * Open a workbook in pieces instead of copying the whole file into memory.
 * Browser: a File/Blob, Web Workers only (throws NOT_IN_WORKER elsewhere).
 * Node: a file path (closed on free()) or an open fd (you close it).
 */
export declare const LazyWorkbook: { open(file: Blob | string | number): Workbook };
`,
);

// wasm-bindgen emits inline_js snippets as ESM even for --target nodejs, and the
// CJS entry require()s them. That breaks wherever require(esm) is missing (Jest,
// Node < 22.12). Rewrite them to CJS.
// ponytail: handles only `export function name`; extend if a snippet exports anything else.
for (const dir of ["pkg/nodejs/snippets", "pkg/writer/nodejs/snippets"]) {
  for (const f of fs.readdirSync(dir, { recursive: true })) {
    if (!f.endsWith(".js")) continue;
    const p = `${dir}/${f}`;
    const src = fs.readFileSync(p, "utf8").replace(/^export function (\w+)/gm, "exports.$1 = function $1");
    if (/^\s*(export|import)\b/m.test(src)) throw new Error(`ESM left in ${p}`);
    fs.writeFileSync(p, src);
  }
}
