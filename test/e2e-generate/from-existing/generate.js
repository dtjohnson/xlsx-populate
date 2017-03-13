"use strict";

module.exports = workbook => {
    workbook.sheet(0).cell("A3").value("baz");
};
