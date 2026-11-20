// LazyWorkbook for Node: read a file by path (closed on free()) or by fd (caller closes it).
const fs = require("node:fs");

module.exports = function lazyWorkbook(Workbook) {
  return {
    open(file) {
      const own = typeof file !== "number";
      const fd = own ? fs.openSync(file, "r") : file;
      let wb;
      try {
        wb = Workbook.openSource({
          size: fs.fstatSync(fd).size,
          readAt(offset, length) {
            const buf = Buffer.allocUnsafe(length);
            let n = 0;
            while (n < length) {
              const got = fs.readSync(fd, buf, n, length - n, offset + n);
              if (got === 0) break;
              n += got;
            }
            return n === length ? buf : buf.subarray(0, n);
          },
        });
      } catch (e) {
        if (own) fs.closeSync(fd);
        throw e;
      }
      if (own) {
        const free = wb.free.bind(wb);
        let closed = false;
        wb.free = wb[Symbol.dispose] = () => {
          free();
          if (!closed) fs.closeSync(fd);
          closed = true;
        };
      }
      return wb;
    },
  };
};
