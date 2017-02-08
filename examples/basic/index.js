"use strict";

const Workbook = require('../../lib/Workbook');

// Load the input workbook from file.
Workbook.fromFileAsync("../shared-formulas/template.xlsx")
    .then(workbook => {
        // Modify the workbook.
        // workbook.sheet("Sheet1").cell("A1").value("This is neat!").value();

        // Write to file.
        return workbook.toFileAsync("./out.xlsx");
    });
