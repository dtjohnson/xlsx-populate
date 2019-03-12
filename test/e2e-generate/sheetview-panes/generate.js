"use strict";

module.exports = workbook => {
    workbook.sheet(0).freezePanes(1, 1);
    workbook.addSheet('sheet2').freezePanes(0, 1);
    workbook.addSheet('sheet3').freezePanes(1, 0);
    workbook.addSheet('sheet4').freezePanes('B2');
    workbook.addSheet('sheet5').freezePanes('B1');
    workbook.addSheet('sheet6').freezePanes('A2');
    workbook.addSheet('sheet7').splitPanes(2000, 2000);
};
