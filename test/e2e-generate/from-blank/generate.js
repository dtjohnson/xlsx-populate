"use strict";

module.exports = workbook => {
    const sheet = workbook.sheet(0);
    sheet.cell("A1").value("TEST");
    sheet.cell("A2").formula("5*2");
};
