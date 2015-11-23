/* jshint jasmine: true */

"use strict";

var Cell = require('../lib/Cell');

describe("Cell", function () {
    describe("getSheet", function () {
        it("should return the parent sheet object", function () {
            var sheet = {};
            var cell = new Cell(sheet, null, null, null);
            expect(cell.getSheet()).toBe(sheet);
        });
    });

    describe("getRow", function () {
        it("should return the row", function () {
            var cell = new Cell(null, 5, null, null);
            expect(cell.getRow()).toBe(5);
        });
    });

    describe("getColumn", function () {
        it("should return the column", function () {
            var cell = new Cell(null, null, 3, null);
            expect(cell.getColumn()).toBe(3);
        });
    });

    describe("getAddress", function () {
        it("should return the address", function () {
            var cell = new Cell(null, 5, 3, null);
            expect(cell.getAddress()).toBe("C5");
        });
    });

    describe("getFullAddress", function () {
        it("should return the full address", function () {
            var sheetMock = {
                getName: function () {
                    return "Foo";
                }
            };

            var cell = new Cell(sheetMock, 5, 3, null);
            expect(cell.getFullAddress()).toBe("'Foo'!C5");
        });
    });

    describe("setValue", function () {
    });

    describe("setFormula", function () {
    });

    describe("_clearContents", function () {
    });
});