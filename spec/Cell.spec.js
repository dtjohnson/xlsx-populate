"use strict";

var xpath = require('xpath');
var DOMParser = require('xmldom').DOMParser;
var parser = new DOMParser();
var Cell = require("../lib/Cell");

describe("Cell", function () {
    var row, sheet, cellNode, cell;

    beforeEach(function () {
        sheet = {
            getName: jasmine.createSpy("sheet.getName").and.returnValue("Foo")
        };
        row = {
            getSheet: jasmine.createSpy("row.getSheet").and.returnValue(sheet)
        };
        cellNode = parser.parseFromString('<c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="C5" t="b"><v>1</v></c>').documentElement;
        cell = new Cell(row, cellNode);
    });

    describe("getSheet", function () {
        it("should return the sheet", function () {
            expect(cell.getSheet()).toBe(sheet);
            expect(row.getSheet).toHaveBeenCalledWith();
        });
    });

    describe("getRow", function () {
        it("should return the row", function () {
            expect(cell.getRow()).toBe(row);
        });
    });

    describe("getAddress", function () {
        it("should return the address", function () {
            expect(cell.getAddress()).toBe("C5");
        });
    });

    describe("getRowNumber", function () {
        it("should return the row number", function () {
            expect(cell.getRowNumber()).toBe(5);
        });
    });

    describe("getColumnNumber", function () {
        it("should return the column number", function () {
            expect(cell.getColumnNumber()).toBe(3);
        });
    });

    describe("getColumnName", function () {
        it("should return the column name", function () {
            expect(cell.getColumnName()).toBe("C");
        });
    });

    describe("getFullAddress", function () {
        it("should return the full address", function () {
            expect(cell.getFullAddress()).toBe("'Foo'!C5");
            expect(sheet.getName).toHaveBeenCalledWith();
        });
    });

    describe("setValue", function () {
        it("should return the cell after setting the value", function () {
            expect(cell.setValue(5)).toBe(cell);
        });

        it("should store a number", function () {
            cell.setValue(57.8);
            expect(cellNode.toString()).toBe('<c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="C5"><v>57.8</v></c>');

            cell.setValue(-6);
            expect(cellNode.toString()).toBe('<c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="C5"><v>-6</v></c>');

            cell.setValue(0);
            expect(cellNode.toString()).toBe('<c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="C5"><v>0</v></c>');
        });

        it("should store a boolean", function () {
            cell.setValue(true);
            expect(cellNode.toString()).toBe('<c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="C5" t="b"><v>1</v></c>');

            cell.setValue(false);
            expect(cellNode.toString()).toBe('<c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="C5" t="b"><v>0</v></c>');
        });

        it("should store a string", function () {
            cell.setValue("some string");
            expect(cellNode.toString()).toBe('<c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="C5" t="inlineStr"><is><t>some string</t></is></c>');
        });

        it("should store a date", function () {
            cell.setValue(new Date('01 Jan 2016 00:00:00'));
            expect(cellNode.toString()).toBe('<c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="C5"><v>42370</v></c>');
        });

        it("should clear the cell if null or undefined", function () {
            cell.setValue(null);
            expect(cellNode.toString()).toBe('<c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="C5"/>');

            cell.setValue();
            expect(cellNode.toString()).toBe('<c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="C5"/>');
        });
    });

    describe("setFormula", function () {
        it("should clear the formula if set to nothing", function () {
            cell.setFormula();
            expect(cellNode.toString()).toBe('<c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="C5"><f/></c>');
        });

        it("should clear the formula if set to empty string", function () {
            cell.setFormula('');
            expect(cellNode.toString()).toBe('<c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="C5"><f/></c>');
        });

        it("should set the formula", function () {
            cell.setFormula('5+6');
            expect(cellNode.toString()).toBe('<c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="C5"><f>5+6</f></c>');
        });

        it("should set the formula with a precalculated value", function () {
            cell.setFormula('ISNUMBER("foo")', false);
            expect(cellNode.toString()).toBe('<c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="C5" t="b"><v>0</v><f>ISNUMBER(\"foo\")</f></c>');
        });
    });

    describe("_clearContents", function () {
        it("should clear the node contents", function () {
            expect(cell._cellNode.childNodes.length).toBe(1);
            expect(cell._cellNode.getAttribute("t")).toBeTruthy();
            cell._clearContents();
            expect(cell._cellNode.childNodes.length).toBe(0);
            expect(cell._cellNode.getAttribute("t")).toBe("");
        });
    });
});
