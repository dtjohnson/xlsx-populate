"use strict";

var jsdoc2md = require("jsdoc-to-markdown");
var fs = require('fs');

// Copy the base README.md
fs.writeFileSync('./README.md', fs.readFileSync('./docs/README.md'));

// Pipe the JSDoc output to the end of the file.
jsdoc2md({ src: "lib/*.js" }).pipe(fs.createWriteStream('./README.md', { flags: 'a' }));
