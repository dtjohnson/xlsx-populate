"use strict";

var path = require('path');

// Load the input workbook from file.
var Workbook = require('../../lib/Workbook');

// Get template workbook and sheet.
var workbook = Workbook.fromFileSync(path.join(__dirname, 'template.xlsx'));
var sheet = workbook.sheet('ClickThroughRateSheet');

// Get header cells.
var clicksHeader = sheet.cell('B2');
var impressionsHeader = sheet.cell('C2');
var ctrHeader = sheet.cell('D2');

// Randomly generate 10 rows of data.
var r = 0;
while (r < 10) {
    r++; // Skip header
    
    var clickValue = parseInt(1e3 * Math.random());
    var impressionValue = parseInt(1e6 * Math.random());

    clicksHeader.relativeCell(r, 0).value(clickValue);
    impressionsHeader.relativeCell(r, 0).value(impressionValue);
}

// Assign shared formulas.
ctrHeader
	.relativeCell(1, 0) // Start from the first cell below header
	.shareFormulaUntil(ctrHeader.relativeCell(r, 0)) // End at the last modifed row
	;

// Write to file.
workbook.toFileSync('./out.xlsx');
