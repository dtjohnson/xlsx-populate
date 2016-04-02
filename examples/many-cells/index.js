"use strict";

var Workbook = require('../../lib/Workbook');

// Load the input workbook from file.
var workbook = Workbook.fromBlankSync();

// For performance, you want to 'get' the objects a little as possible and hold onto references.
var sheet = workbook.getSheet("Sheet1");
for (var rowNumber = 1; rowNumber <= 1000; rowNumber++) {
    var row = sheet.getRow(rowNumber);
    for (var columnNumber = 1; columnNumber <= 100; columnNumber++) {
        row.getCell(columnNumber).setValue("foo");
    }
}

// Write to file.
workbook.toFileSync("./out.xlsx");
