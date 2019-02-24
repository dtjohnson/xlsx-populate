"use strict";

const { Workbook } = require('../../dist/');

(async () => {
    // Load the input workbook from file.
    const workbook = await Workbook.fromBlankAsync();

    // Modify the workbook.
    workbook.sheet("Sheet1").cell("A1").value("This is neat!");

    // Write to file.
    await workbook.toFileAsync("./out.xlsx");
})();
