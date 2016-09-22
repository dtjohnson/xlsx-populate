"use strict";

var path = require('path');

// Load the input workbook from file.
var Workbook = require('../../lib/Workbook');

// Get template workbook and sheet.
var workbook = Workbook.fromFileSync(path.join(__dirname, 'template.xlsx'));
var sheet = workbook.getSheet('ClickThroughRateSheet');

// Get header cells.
var clicksHeader = sheet.getCell('B2');
var impressionsHeader = sheet.getCell('C2');
var ctrHeader = sheet.getCell('D2');

// Randomly generate 10 rows of data.
var r = 0;
while (r < 10) {
    r++; // Skip header
    
    var clickValue = parseInt(1e3 * Math.random());
    var impressionValue = parseInt(1e6 * Math.random());

    clicksHeader.getRelativeCell(r, 0).setValue(clickValue);
    impressionsHeader.getRelativeCell(r, 0).setValue(impressionValue);
}

// Assign shared formulas.
ctrHeader
	.getRelativeCell(1, 0) // Start from the first cell below header
	.shareFormulaUntil(ctrHeader.getRelativeCell(r, 0)) // End at the last modifed row
	;

// Write to file.
workbook.toFileSync('./out.xlsx');
