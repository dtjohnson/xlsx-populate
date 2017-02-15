"use strict";

const Workbook = require('../../lib/Workbook');

// Load the input workbook from file.
Workbook.fromBlankAsync()
    .then(workbook => {
        // Modify the workbook.
        const r = workbook.sheet("Sheet1").range("A1:B2").value(() => Math.random());


        // workbook.sheet("Sheet1").row(3).hidden(true).hidden(false);

        // Write to file.
        return workbook.toFileAsync("./out.xlsx");
    });
