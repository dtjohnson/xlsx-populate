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
            sheet: jasmine.createSpy("row.getSheet").and.returnValue(sheet)
        };
        cellNode = parser.parseFromString('<c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="C5" t="b"><v>1</v></c>').documentElement;
        cell = new Cell(row, cellNode);
    });

    describe("sheet", function () {
        it("should return the sheet", function () {
            expect(cell.sheet()).toBe(sheet);
            expect(row.sheet).toHaveBeenCalledWith();
        });
    });

    describe("row", function () {
        it("should return the row", function () {
            expect(cell.row()).toBe(row);
        });
    });

    describe("address", function () {
        it("should return the address", function () {
            expect(cell.address()).toBe("C5");
        });
    });

    describe("rowNumber", function () {
        it("should return the row number", function () {
            expect(cell.rowNumber()).toBe(5);
        });
    });

    describe("columnNumber", function () {
        it("should return the column number", function () {
            expect(cell.columnNumber()).toBe(3);
        });
    });

    describe("columnName", function () {
        it("should return the column name", function () {
            expect(cell.columnName()).toBe("C");
        });
    });

    describe("fullAddress", function () {
        it("should return the full address", function () {
            expect(cell.fullAddress()).toBe("'Foo'!C5");
            expect(sheet.getName).toHaveBeenCalledWith();
        });
    });

    describe("value", function () {
        it("should return the cell after setting the value", function () {
            expect(cell.value(5)).toBe(cell);
        });

        it("should store a number", function () {
            cell.value(57.8);
            expect(cellNode.toString()).toBe('<c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="C5"><v>57.8</v></c>');

            cell.value(-6);
            expect(cellNode.toString()).toBe('<c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="C5"><v>-6</v></c>');

            cell.value(0);
            expect(cellNode.toString()).toBe('<c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="C5"><v>0</v></c>');
        });

        it("should store a boolean", function () {
            cell.value(true);
            expect(cellNode.toString()).toBe('<c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="C5" t="b"><v>1</v></c>');

            cell.value(false);
            expect(cellNode.toString()).toBe('<c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="C5" t="b"><v>0</v></c>');
        });

        it("should store a string", function () {
            cell.value("some string");
            expect(cellNode.toString()).toBe('<c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="C5" t="inlineStr"><is><t>some string</t></is></c>');
        });

        it("should store a date", function () {
            cell.value(new Date('01 Jan 2016 00:00:00'));
            expect(cellNode.toString()).toBe('<c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="C5"><v>42370</v></c>');
        });

        it("should clear the cell if null or undefined", function () {
            cell.value(null);
            expect(cellNode.toString()).toBe('<c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="C5"/>');

            cell.value();
            expect(cellNode.toString()).toBe('<c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="C5"/>');
        });
    });

    describe("formula", function () {
        it("should clear the formula if set to nothing", function () {
            cell.formula();
            expect(cellNode.toString()).toBe('<c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="C5"><f/></c>');
        });

        it("should clear the formula if set to empty string", function () {
            cell.formula('');
            expect(cellNode.toString()).toBe('<c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="C5"><f/></c>');
        });

        it("should set the formula", function () {
            cell.formula('5+6');
            expect(cellNode.toString()).toBe('<c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="C5"><f>5+6</f></c>');
        });

        it("should set the formula with a precalculated value", function () {
            cell.formula('ISNUMBER("foo")', false);
            expect(cellNode.toString()).toBe('<c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="C5" t="b"><v>0</v><f>ISNUMBER(\"foo\")</f></c>');
        });
    });

    describe("clear", function () {
        it("should clear the node contents", function () {
            expect(cell._cellNode.childNodes.length).toBe(1);
            expect(cell._cellNode.getAttribute("t")).toBeTruthy();
            cell.clear();
            expect(cell._cellNode.childNodes.length).toBe(0);
            expect(cell._cellNode.getAttribute("t")).toBe("");
        });
    });
});
