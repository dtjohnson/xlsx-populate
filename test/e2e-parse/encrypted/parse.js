"use strict";

module.exports = workbook => {
    return workbook.sheet(0).cell("A1").value();
};
