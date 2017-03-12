"use strict";

require("../../spec/helpers/async");
require("../../spec/helpers/toEqualJson");
const foo = require("./foo");
const XlsxPopulate = require("../../lib/XlsxPopulate");

describe("test", () => {
    itAsync("should foo", () => {
        return XlsxPopulate.fromBlankAsync()
            .then(workbook => {
                workbook.sheet(0).cell("A1").value("TEST");
                return workbook.toFileAsync("./out.xlsx");
            })
            .then(() => foo("./foo.cs"))
            .then(result => {
                expect(result).toEqualJson({ foo: "TEST" });
            });
    });
});
