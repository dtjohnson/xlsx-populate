"use strict";

var proxyquire = require("proxyquire").noCallThru();
var xpath = require('xpath');
var DOMParser = require('xmldom').DOMParser;
var parser = new DOMParser();

describe("Row", function () {
    var Cell, Row, sheet, rowNode, row;

    beforeEach(function () {
        Cell = jasmine.createSpy("Cell");
        Row = proxyquire("../lib/Row", { './Cell': Cell });
        sheet = {};
        rowNode = parser.parseFromString('<row r="7"><c r="A1"/><c r="K7"/></row>').documentElement;
        row = new Row(sheet, rowNode);
    });

    describe("getSheet", function () {
        it("should return the sheet", function () {
            expect(row.getSheet()).toBe(sheet);
        });
    });

    describe("getRowNumber", function () {
        it("should return the row number", function () {
            expect(row.getRowNumber()).toBe(7);
        });
    });

    describe("getCell", function () {
        it("should create a new cell node if it doesn't exist", function () {
            row.getCell(7);
            expect(Cell).toHaveBeenCalledWith(row, rowNode.lastChild);
            expect(rowNode.toString()).toBe('<row r="7"><c r="A1"/><c r="K7"/><c r="G7"/></row>');
        });

        it("should use an existing cell node if it does exist", function () {
            row.getCell(11);
            expect(Cell).toHaveBeenCalledWith(row, rowNode.lastChild);
            expect(rowNode.toString()).toBe('<row r="7"><c r="A1"/><c r="K7"/></row>');
        });
    });
});
