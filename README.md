# msoconvert

Convert Word and Excel documents in Node.js (for Windows only). This is a wrapper around [DocTo](https://github.com/tobya/DocTo).

Microsoft Word/Excel must be installed on the system.

## Installation

```
npm install msoconvert --save
```

## Usage

```javascript
const convert = require("msoconvert");

// All the folling options are strings:
convert({ 
  use, // "word" (default) | "excel" | "powerpoint"
  input, // Input file or directory path.
  output, // Output file or directory path.
  format, // wdSaveFormat for output (see below).

  // The two following only apply if input is a directory:
  inputExt, // Extension to search for if directory.
  outputExt, // Output extension.
  options // Additional arguments passed to docto.exe. Refer to DocTo documentation.
})
  // A promise is returned.
  .then(() => console.log("done"));
```

See [DocTo documentation](https://github.com/tobya/DocTo) for more information about API and additional parameters.

### `format`

`wdSaveFormat` enums:
* https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.wdsaveformat
* https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.xlfileformat

If `format` is not defined, msoconvert will try to guess the format from `outputExt` if provided (otherwise an error will be thrown).