"use strict";

const XlsxPopulate = require('../../lib/XlsxPopulate');
const JSZip = require('JSZip');
const fs = require('fs');

// Load the input workbook from file.
XlsxPopulate.fromFileAsync("in.xlsx").then(workbook => {
    return workbook.toFileAsync('out.xlsx')
})
.catch(err => console.error(err));
