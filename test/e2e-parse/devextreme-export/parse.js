"use strict";

module.exports = workbook => {
    return workbook.sheet(0).usedRange().value();
};
