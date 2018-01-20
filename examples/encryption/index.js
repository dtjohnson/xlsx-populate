"use strict";

const XlsxPopulate = require('../../lib/XlsxPopulate');

// Load the input workbook from file with a password.
XlsxPopulate.fromFileAsync("./in.xlsx", { password: "open sesame" })
    .then(workbook => {
        // Read a value from the workbook.
        console.log(workbook.sheet("Sheet1").cell("A1").value());

        // Write to file with a new password.
        return workbook.toFileAsync("./out.xlsx", { password: "new password" });
    })
    .catch(err => console.error(err));
