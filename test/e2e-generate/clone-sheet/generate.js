"use strict";

module.exports = workbook => {
    workbook
        .cloneSheet(workbook.sheet(0), 'Sheet Cloned')
        .cell("A3").value("baz");
};
