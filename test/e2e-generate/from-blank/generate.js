"use strict";

module.exports = workbook => {
    workbook.sheet(0).cell("A1").value("TEST");
};
