const test = require("ava");
const fs = require("fs-extra");
const path = require("path");
const convert = require("./index.js");

const fileExists = (s) => new Promise(r => fs.access(s, fs.F_OK, e => r(!e)));

const allFilesExist = (f) => Promise.all(f.map(fileExists)).then(r => r.every(v => v === true));

test.beforeEach(async () => {
  const outDir = path.join(__dirname, "test/out");
  return fs.remove(outDir);
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

test.serial("Define output encoding", async t => {
  const options = {
    input: "test\\src\\test.docx",
    output: "test\\out\\test.html",
    format: "wdFormatHTML",
    encoding: "UTF-8"
  };
  await convert(options);
  const html = await fs.readFile(options.output);
  const isUTF8 = html.includes("<meta http-equiv=Content-Type content=\"text/html; charset=utf-8\">");
  t.true(isUTF8);
});