"use strict";

const XlsxPopulate = require('../../lib/XlsxPopulate');

// Load the input workbook from file.
XlsxPopulate.fromBlankAsync()
    .then(workbook => {
        // Modify the workbook.
        workbook.sheet("Sheet1").cell("A1").value("This is neat!");

        // Write to file.
        return workbook.toFileAsync("./out.xlsx");
    })
    .catch(err => console.error(err));
