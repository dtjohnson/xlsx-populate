"use strict";

/* eslint no-console:off */

// Load the input workbook from file.
const XlsxPopulate = require('../../lib/XlsxPopulate');

// Get template workbook and sheet.
XlsxPopulate.fromFileAsync('./template.xlsx')
    .then(workbook => {
        // Randomly generate 10 rows of data.
        const sheet = workbook.sheet('ClickThroughRateSheet');
        sheet.range("B3:B13").value(() => parseInt(1e3 * Math.random()));
        sheet.range("C3:C13").value(() => parseInt(1e6 * Math.random()));
        sheet.range("D3:D13").formula("B3/C3").style("numberFormat", "0.00%");

        console.log(sheet.usedRange().value());

        // Write to file.
        return workbook.toFileAsync('./out.xlsx');
    })
    .catch(err => console.error(err));
