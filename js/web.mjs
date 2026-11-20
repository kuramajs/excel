// Browser entry: the wasm-bindgen build, with an init() that explains the most common deploy failure.
import wasmInit, { Workbook } from "../pkg/web/web_index.js";
import { withHelp } from "./with-help.mjs";

export * from "../pkg/web/web_index.js";
export default withHelp(wasmInit, "web_index_bg.wasm");

export const LazyWorkbook = {
  /** Open a File/Blob in pieces. Web Workers only: reads must be synchronous (FileReaderSync). */
  open(blob) {
    if (typeof FileReaderSync === "undefined") {
      const err = new Error(
        "LazyWorkbook.open needs a Web Worker (FileReaderSync). On the main thread use " +
          "Workbook.open(await blob.arrayBuffer()), or move parsing into a worker.",
      );
      err.code = "NOT_IN_WORKER";
      throw err;
    }
    if (!(blob instanceof Blob)) {
      const err = new Error("LazyWorkbook.open expects a File or Blob");
      err.code = "INVALID_INPUT";
      throw err;
    }
    const reader = new FileReaderSync();
    return Workbook.openSource({
      size: blob.size,
      readAt: (offset, length) => reader.readAsArrayBuffer(blob.slice(offset, offset + length)),
    });
  },
};
