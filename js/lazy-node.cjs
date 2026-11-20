// LazyWorkbook for Node: read a file by path (closed on free(), or on garbage collection if
// free() is never called) or by fd (caller closes it).
const fs = require("node:fs");

const close = (file) => {
  if (!file.closed) fs.closeSync(file.fd);
  file.closed = true;
};
// Keyed on readAt, not the workbook: the wasm side holds readAt until the workbook and every
// iterator borrowing its reader are gone, so the file can't close under a running iterRows().
const unfreed = new FinalizationRegistry(close);

module.exports = function lazyWorkbook(Workbook) {
  return {
    open(file) {
      const own = typeof file !== "number";
      const fd = own ? fs.openSync(file, "r") : file;
      const readAt = (offset, length) => {
        const buf = Buffer.allocUnsafe(length);
        let n = 0;
        while (n < length) {
          const got = fs.readSync(fd, buf, n, length - n, offset + n);
          if (got === 0) break;
          n += got;
        }
        return n === length ? buf : buf.subarray(0, n);
      };
      let wb;
      try {
        wb = Workbook.openSource({ size: fs.fstatSync(fd).size, readAt });
      } catch (e) {
        if (own) fs.closeSync(fd);
        throw e;
      }
      if (own) {
        const state = { fd, closed: false };
        unfreed.register(readAt, state, state);
        const free = wb.free.bind(wb);
        wb.free = wb[Symbol.dispose] = () => {
          free();
          unfreed.unregister(state);
          close(state);
        };
      }
      return wb;
    },
  };
};
