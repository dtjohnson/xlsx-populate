"use strict";

var Workbook = require('../../lib/Workbook');

// Load the input workbook from file.
Workbook.fromBlankAsync()
    .then(workbook => {
        // Modify the workbook.
        workbook.sheet("Sheet1").cell("A1").value("This is neat!");

        // Write to file.
        workbook.toFileAsync("./out.xlsx");
    })
    .catch(err => console.error(err.stack));
