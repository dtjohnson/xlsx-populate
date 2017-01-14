"use strict";

const proxyquire = require("proxyquire").noCallThru();
const xpath = require('xpath');
const DOMParser = require('xmldom').DOMParser;
const parser = new DOMParser();

describe("Cell", () => {
    let Cell, utils, row, sheet, cellNode, cell;

    beforeEach(() => {
        utils = jasmine.createSpyObj("utils", ["addressToRowAndColumn", "columnNumberToName"]);
        utils.addressToRowAndColumn.and.returnValue({ row: "ROW", column: "COLUMN" });
        utils.columnNumberToName.and.returnValue("NAME");

        Cell = proxyquire("../lib/Cell", {
            './utils': utils
        });

        sheet = {
            name: jasmine.createSpy("sheet.name").and.returnValue("SHEET_NAME")
        };
        row = {
            sheet: jasmine.createSpy("row.sheet").and.returnValue(sheet),
            workbook: jasmine.createSpy("sheet.name").and.returnValue("WORKBOOK")
        };
        cellNode = parser.parseFromString('<c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="C5" t="b"><v>1</v></c>').documentElement;
        cell = new Cell(row, cellNode);
    });

    describe("constructor", () => {
        it("should save the row and cell node", () => {
            expect(cell._row).toBe(row);
            expect(cell._cellNode).toBe(cellNode);
        });
    });

    describe("address", () => {
        it("should return the address", () => {
            expect(cell.address()).toBe("C5");
        });

        it("should throw an error if a value is provided", () => {
            expect(() => cell.address("foo")).toThrow();
        });
    });

    describe("columnName", () => {
        it("should return the column name", () => {
            expect(cell.columnName()).toBe("NAME");
            expect(utils.columnNumberToName).toHaveBeenCalledWith("COLUMN");
        });

        it("should throw an error if a value is provided", () => {
            expect(() => cell.columnName("foo")).toThrow();
        });
    });

    describe("columnNumber", () => {
        it("should return the column number", () => {
            expect(cell.columnNumber()).toBe("COLUMN");
            expect(utils.addressToRowAndColumn).toHaveBeenCalledWith("C5");
        });

        it("should throw an error if a value is provided", () => {
            expect(() => cell.columnNumber("foo")).toThrow();
        });
    });

    describe("row", () => {
        it("should return the row", () => {
            expect(cell.row()).toBe(row);
        });

        it("should throw an error if a value is provided", () => {
            expect(() => cell.row("foo")).toThrow();
        });
    });

    describe("rowNumber", () => {
        it("should return the row number", () => {
            expect(cell.rowNumber()).toBe("ROW");
            expect(utils.addressToRowAndColumn).toHaveBeenCalledWith("C5");
        });

        it("should throw an error if a value is provided", () => {
            expect(() => cell.rowNumber("foo")).toThrow();
        });
    });

    describe("sheet", () => {
        it("should return the sheet", () => {
            expect(cell.sheet()).toBe(sheet);
            expect(row.sheet).toHaveBeenCalledWith();
        });

        it("should throw an error if a value is provided", () => {
            expect(() => cell.sheet("foo")).toThrow();
        });
    });

    describe("workbook", () => {
        it("should return the sheet", () => {
            expect(cell.workbook()).toBe("WORKBOOK");
            expect(row.workbook).toHaveBeenCalledWith();

            it("should throw an error if a value is provided", () => {
                expect(() => cell.workbook("foo")).toThrow();
            });
        });
    });











    xdescribe("fullAddress", function () {
        it("should return the full address", function () {
            expect(cell.fullAddress()).toBe("'Foo'!C5");
            expect(sheet.getName).toHaveBeenCalledWith();
        });
    });

    xdescribe("value", function () {
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

    xdescribe("formula", function () {
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

    xdescribe("clear", function () {
        it("should clear the node contents", function () {
            expect(cell._cellNode.childNodes.length).toBe(1);
            expect(cell._cellNode.getAttribute("t")).toBeTruthy();
            cell.clear();
            expect(cell._cellNode.childNodes.length).toBe(0);
            expect(cell._cellNode.getAttribute("t")).toBe("");
        });
    });
});
