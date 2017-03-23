"use strict";

module.exports = workbook => {
    const sheet1 = workbook.sheet(0);
    const sheet2 = workbook.addSheet("Sheet2");
    const sheet3 = workbook.addSheet("Sheet3");

    workbook.moveSheet(sheet1);
    sheet3.move("Sheet2");

    sheet2.tabColor(3);
    sheet3.active(true);
    sheet1.tabSelected(true);
};
