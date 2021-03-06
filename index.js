const exec = require("child_process").exec;
const path = require("path");

const defaultFormats = require("./enum/default-formats.json");
const encodings = require("./enum/msoencoding.json");

const cwd = process.cwd();
const exePath = path.join(__dirname, "bin/docto.exe");

const getAbsolutePath = (p) => path.isAbsolute(p) ? p : path.join(cwd, p);

function getUseArg(use) {
  return (use === "excel" || use === "powerpoint") ? "--" + use : "--word";
}

function getOutputExtArg(outputExt) {
  if (outputExt == null) return "";
  if (outputExt.charAt(0) !== ".") {
    outputExt = "." + outputExt;
  }
  return "--outputextension " + outputExt;
}

function getFormatArg(format, outputExt) {
  if (format) return `--format ${format}`;
  const ext = (outputExt.match(/[A-z]*$/) || [])[0];
  const guess = defaultFormats[ext];
  if (!guess) {
    throw Error("Could not guess format.");
  }
  return `--format ${guess}`;
}

function getEncodingArg(encoding) {
  if (!encoding) return "";
  if (Number(encoding) === encoding) return `-E ${encoding}`;
  const key = encoding.toLowerCase().replace(/[^A-z0-9]+/gm, "");
  const guess = encodings[key] || encodings["msoencoding" + key];
  if (!guess) {
    throw Error("Unknown encoding.");
  }
  return `-E ${guess}`;
}

function run(args) {
  return new Promise((resolve, reject) => {
    const command = `"${exePath}" ${args.join(" ")}`;
    return exec(command, (error, stdout, stderr) => {
      if (error) {
        reject(error);
      }
      resolve(stdout? stdout : stderr);
    });
  });
}

function convert({ use, input, inputExt, output, outputExt, format, encoding, options = "" }) {
  if (input == null || output == null) {
    return Promise.reject("Missing parameter: input and output must be defined.");
  }
  if (format == null && outputExt == null) {
    return Promise.reject("Missing parameter: format or outputExt must be defined.");
  }

  const args = [
    getUseArg(use),
    `--inputfile "${input}"`,
    inputExt ? `-FX ${inputExt}` : "",
    `--outputfile "${getAbsolutePath(output)}"`,
    getOutputExtArg(outputExt),
    getFormatArg(format, outputExt),
    getEncodingArg(encoding),
    options
  ];
  
  return run(args);
}

module.exports = convert;