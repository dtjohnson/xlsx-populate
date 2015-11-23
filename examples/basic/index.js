"use strict";

var Workbook = require('../../lib/Workbook');

// Load the input workbook from file.
var workbook = Workbook.fromBlankSync();

// Modify the workbook.
workbook.getSheet("Sheet1").getCell("A1").setValue("This is neat!");

// Write to file.
workbook.toFileSync("./out.xlsx");
