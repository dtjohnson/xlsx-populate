"use strict";

const XlsxPopulate = require('../../lib/XlsxPopulate');
const JSZip = require('JSZip');
const fs = require('fs');

// Load the input workbook from file.
XlsxPopulate.fromFileAsync("in.xlsx").then(workbook => {
    workbook.sheet(0).drawings('Test Image').name('new_test_image')
    workbook.sheet(3).drawings('Picture 2').name('new_test_image_2')
    return workbook.toFileAsync('out.xlsx')
})
.catch(err => console.error(err));