"use strict";

module.exports = workbook => {
    const sheet = workbook.sheet(0);
    sheet.pageMargins('left', 0.1);
    sheet.pageMargins('right', 0.2);
    sheet.pageMargins('top', 0.3);
    sheet.pageMargins('bottom', 0.4);
    sheet.pageMargins('header', 0.5);
    sheet.pageMargins('footer', 0.6);
};
