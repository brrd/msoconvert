const util = require("util");
const exec = util.promisify(require("child_process").exec);
const path = require("path");

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

function convert({ use, input, inputExt, output, outputExt, format, options = "" }) {
  if (input == null || output == null || format == null) {
    return Promise.reject("Missing parameter (input, output, format).");
  }

  const args = [
    getUseArg(use),
    `--inputfile "${input}"`,
    inputExt ? `-FX ${inputExt}` : "",
    `--outputfile "${getAbsolutePath(output)}"`,
    getOutputExtArg(outputExt),
    `--format ${format}`,
    options
  ];
  
  return run(args);
}

module.exports = convert;