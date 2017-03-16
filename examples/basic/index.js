"use strict";

const XlsxPopulate = require('../../lib/XlsxPopulate');

// Load the input workbook from file.
XlsxPopulate.fromFileAsync("../Book1.xlsx")
    .then(workbook => {
        // Modify the workbook.

        // workbook.moveSheet("Sheet1");
        // workbook.addSheet("NEW").tabColor("0000FF").active(true);
        console.log(workbook.find(2));
        // workbook.activeSheet().tabColor(6).activeCell().value("FOO");
        // console.log(workbook.activeSheet().tabColor())

        // Write to file.
        return workbook.toFileAsync("./out.xlsx");
    }).catch(err => console.error(err));
