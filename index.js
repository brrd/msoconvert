const util = require("util");
const exec = util.promisify(require("child_process").exec);
const path = require("path");

const exePath = path.join(__dirname, "bin/docto.exe");

function getUseArg(use) {
  return (use === "excel" || use === "powerpoint") ? "--" + use : "--word";
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

function file({ use, input, output, format, options = "" }) {
  if (input == null || output == null || format == null) {
    return Promise.reject("Missing parameter (input, output or format).");
  }

  const args = [
    getUseArg(use),
    `--inputfile "${input}"`,
    `--outputfile "${output}"`,
    `--format ${format}`,
    options
  ];
  
  return run(args);
}

module.exports = { file };