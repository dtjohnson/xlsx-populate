"use strict";

const XlsxPopulate = require('../../lib/XlsxPopulate');
const JSZip = require('JSZip');
const fs = require('fs');

// Load the input workbook from file.
XlsxPopulate.fromFileAsync("in.xlsx").then(workbook => {
    var tmpzip = new JSZip();
    tmpzip.file('image', fs.readFileSync('new_image.png', {encoding: 'binary'})).generateAsync({type : "string"}).then(() => {
        workbook.swapImage('test_image', tmpzip.files.image._data);
        workbook.toFileAsync("./out.xlsx")
    })
})
.catch(err => console.error(err));
