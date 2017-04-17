"use strict";

module.exports = workbook => {
    workbook.sheets().forEach(sheet => {
        sheet.cell("A1").value("FOO");
    });
};
