"use strict";

module.exports = workbook => {
    workbook.sheet(0).freezePanes(1, 1);
    // workbook.addSheet('sheet2').freezePanes('B2');
    // workbook.addSheet('sheet3').splitPanes(2000, 2000);
};
