"use strict";

var Workbook = require('../../lib/Workbook');

// Load the input workbook from file.
Workbook.fromBlankAsync()
    .then(workbook => {
        // For performance, you want to 'get' the objects a little as possible and hold onto references.
        var sheet = workbook.sheet("Sheet1");
        for (var rowNumber = 1; rowNumber <= 1000; rowNumber++) {
            var row = sheet.row(rowNumber);
            for (var columnNumber = 1; columnNumber <= 100; columnNumber++) {
                row.cell(columnNumber).value("foo");
            }
        }

        // Write to file.
        workbook.toFileAsync("./out.xlsx");
    });
