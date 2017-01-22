"use strict";

// Load the input workbook from file.
const Workbook = require('../../lib/Workbook');

// Get template workbook and sheet.
Workbook.fromFileAsync('./template.xlsx')
    .then(workbook => {
        // Randomly generate 10 rows of data.
        const sheet = workbook.sheet('ClickThroughRateSheet');
        sheet.range("B3:B13").forEach(cell => cell.value(parseInt(1e3 * Math.random())));
        sheet.range("C3:C13").forEach(cell => cell.value(parseInt(1e6 * Math.random())));
        sheet.range("D3:D13").formula("B3/C3");

        // Write to file.
        return workbook.toFileAsync('./out.xlsx');
    });
