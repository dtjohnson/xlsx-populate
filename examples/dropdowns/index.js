"use strict";

const XlsxPopulate = require('../../lib/XlsxPopulate');
// Load the input workbook from file.
XlsxPopulate.fromFileAsync("in.xlsx").then(workbook => {
    workbook.find('test1', 'test is working');
    let found = workbook.find('test is working');
    found.forEach(tmp => {
        console.log(tmp.address({includeSheetName: true}));
    })
    return workbook.toFileAsync('out.xlsx')
})
.catch(err => console.error(err));