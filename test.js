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

const allFilesExist = (f) => Promise.all(f.map(fileExists)).then(r => r.every(v => v === true));

test.beforeEach(() => {
  const outDir = path.join(__dirname, "test/out");
  return emptyDir(outDir);
});

test.serial("Convert a single document", async t => {
  const options = {
    input: "test\\src\\test.docx",
    output: "test\\out\\test.doc",
    format: "wdFormatDocument97"
  };
  await convert(options);
  const exists = await fileExists(options.output);
  t.true(exists);
});

test.serial("Guess format from outputExt", async t => {
  const options = {
    input: "test\\src\\test.docx",
    output: "test\\out\\test.doc",
    outputExt: "doc"
  };
  await convert(options);
  const exists = await fileExists(options.output);
  t.true(exists);
});

test.serial("Convert directory", async t => {
  const options = {
    input: "./test/src",
    inputExt: "\"*.docx\"",
    output: "./test/out",
    outputExt: "doc",
    format: "wdFormatDocument97"
  };
  await convert(options);
  const exists = await allFilesExist([
    "test\\out\\test.doc",
    "test\\out\\test2.doc"
  ]);
  t.true(exists);
});