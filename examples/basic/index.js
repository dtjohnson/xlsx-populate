"use strict";

const Workbook = require('../../lib/Workbook');

// Load the input workbook from file.
Workbook.fromBlankAsync()
    .then(workbook => {
        // Modify the workbook.
        workbook.sheet("Sheet1").cell("A1").value("This is neat!");

        // Write to file.
        return workbook.toFileAsync("./out.xlsx");
    });
