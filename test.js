const test = require("ava");
const fs = require("fs");
const path = require("path");
const convert = require("./index.js");

const emptyDir = (dirpath) => {
  return new Promise((resolve, reject) => {
    fs.readdir(dirpath, (err, files) => {
      if (err) reject(err);
      for (const file of files) {
        try {
          if (file === ".keep") continue;
          fs.unlinkSync(path.join(dirpath, file));
        } catch(err) {
          reject(err);
        }
      }
      resolve();
    });
  });
};

const fileExists = (s) => new Promise(r => fs.access(s, fs.F_OK, e => r(!e)));

test.beforeEach(() => {
  const outDir = path.join(__dirname, "test/out");
  return emptyDir(outDir);
});

test.serial("Convert a single DOCX document", async t => {
  const options = {
    input: "test/src/test.docx",
    output: "test/out/test.html",
    format: "wdFormatHTML"
  };
  await convert.file(options);
  const exists = await fileExists(options.output);
  t.true(exists);
});
