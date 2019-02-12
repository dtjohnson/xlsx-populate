"use strict";

module.exports = workbook => {
    const sheet = workbook.sheet(0);
    sheet.printOptions('headings', false);
    sheet.printOptions('horizontalCentered', false);
    sheet.printOptions('verticalCentered', true);
    sheet.printGridLines(false);
};
