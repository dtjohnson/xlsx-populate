"use strict";

var Workbook = require('../../lib/Workbook');

// Load the input workbook from file.
Workbook.fromBlankAsync()
    .then(workbook => {
        // Modify the workbook.
        workbook.getSheet("Sheet1").getCell("A1").setValue("This is neat!");

        // Write to file.
        workbook.toFileAsync("./out.xlsx");
    })
    .catch(err => console.error(err.stack));
