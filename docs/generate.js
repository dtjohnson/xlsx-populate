"use strict";

var jsdoc2md = require("jsdoc-to-markdown");
var replaceStream = require("replacestream");
var fs = require('fs');

// Copy the base README.md
fs.writeFileSync('./README.md', fs.readFileSync('./docs/README.md'));

// Pipe the JSDoc output to the end of the file.
jsdoc2md({ src: "lib/*.js" })
    .pipe(replaceStream(/\* \[new[\S\s]+?\*/g, '*'))// Strip out the constructor definitions since they are private.
    .pipe(replaceStream(/### new[\S\s]+?###/g, '###'))
    .pipe(fs.createWriteStream('./README.md', { flags: 'a' }));
