"use strict";

module.exports = workbook => {
    const range = workbook.sheet(0).usedRange();
    return {
        value: range.value(),
        numberFormat: range.style("numberFormat")
    };
};
