"use strict";

const XlsxPopulate = require('../../lib/XlsxPopulate');

// Load the input workbook from file.
XlsxPopulate.fromFileAsync("../Book1.xlsx")
    .then(workbook => {
        // Modify the workbook.

            console.log(workbook.sheet(4).selection().map(s => s.address()));
        // workbook.activeSheet().cell("A1").value("This is neat!");
        // workbook.activeSheet().activeCell().value("FOO");
    // workbook.activeSheet(2).activeCell("C3");
    //         workbook.sheet(0).active(true);
        // console.log(workbook.sheet(0).hidden());
        // workbook.sheet(0).hidden('true');
        // console.log(workbook.sheet(0).hidden());

        // Write to file.
        return workbook.toFileAsync("./out.xlsx");
    });
