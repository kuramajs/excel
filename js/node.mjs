// Node ESM entry: same API as the web build. The Node build loads its wasm
// synchronously on import, so init() has nothing to do.
import cjs from "./node.cjs";

/** Open a file path (closed on free()) or an open fd in pieces instead of reading it all. */
export const { Workbook, Sheet, RowIter, LazyWorkbook } = cjs;
export function initSync() {}
export default async function init() {}
