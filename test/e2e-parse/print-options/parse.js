"use strict";

module.exports = workbook => {
    const sheet = workbook.sheet(0);
    return {
        printOptions_gridLines: sheet.printOptions('gridLines'),
        printOptions_gridLinesSet: sheet.printOptions('gridLinesSet'),
        printOptions_headings: sheet.printOptions('headings'),
        printOptions_horizontalCentered: sheet.printOptions('horizontalCentered'),
        printOptions_verticalCentered: sheet.printOptions('verticalCentered'),

        printGridLines: sheet.printGridLines()
    };
};
