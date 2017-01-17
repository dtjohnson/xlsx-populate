"use strict";

var Workbook = require('../../lib/Workbook');

// Load the input workbook from file.
Workbook.fromBlankAsync()
    .then(workbook => {
        // For performance, you want to 'get' the objects a little as possible and hold onto references.
        workbook.sheet("Sheet1").range("A1:CZ1000").values("foo");

        // Write to file.
        return workbook.toFileAsync("./out.xlsx");
    });
