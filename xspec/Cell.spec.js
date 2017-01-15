"use strict";

const proxyquire = require("proxyquire").noCallThru();
const xpath = require('xpath');
const DOMParser = require('xmldom').DOMParser;
const parser = new DOMParser();

describe("Cell", () => {
    let Cell, utils, row, sheet, cellNode, cell;

    beforeEach(() => {
        utils = jasmine.createSpyObj("utils", ["addressToRowAndColumn", "columnNumberToName", "addressToFullAddress", "dateToExcelNumber"]);
        utils.addressToRowAndColumn.and.returnValue({ row: "ROW", column: "COLUMN" });
        utils.columnNumberToName.and.returnValue("NAME");
        utils.addressToFullAddress.and.returnValue("FULL_ADDRESS");
        utils.dateToExcelNumber.and.returnValue("EXCEL_NUMBER");

        Cell = proxyquire("../lib/Cell", {
            './utils': utils
        });

        sheet = jasmine.createSpyObj("sheet", ["name", "cell"]);
        sheet.name.and.returnValue("SHEET_NAME");
        sheet.cell.and.returnValue("CELL");

        row = jasmine.createSpyObj("row", ["sheet", "workbook"]);
        row.sheet.and.returnValue(sheet);
        row.workbook.and.returnValue("WORKBOOK");

        cellNode = parser.parseFromString('<c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="C5" t="b"><v>1</v></c>').documentElement;
        cell = new Cell(row, cellNode);
    });

    describe("constructor", () => {
        it("should save the row and cell node", () => {
            expect(cell._row).toBe(row);
            expect(cell._node).toBe(cellNode);
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

    describe("clear", () => {
        it("should clear the node contents", () => {
            expect(cell._node.childNodes.length).toBe(1);
            expect(cell._node.getAttribute("t")).toBeTruthy();
            expect(cell.clear()).toBe(cell);
            expect(cell._node.childNodes.length).toBe(0);
            expect(cell._node.getAttribute("t")).toBe("");
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

    describe("fullAddress", () => {
        it("should return the full address", () => {
            expect(cell.fullAddress()).toBe("FULL_ADDRESS");
            expect(utils.addressToFullAddress).toHaveBeenCalledWith("SHEET_NAME", "C5");
        });

        it("should throw an error if a value is provided", () => {
            expect(() => cell.fullAddress("foo")).toThrow();
        });
    });

    describe("relativeCell", () => {
        beforeEach(() => {
            spyOn(cell, "rowNumber").and.returnValue(5);
            spyOn(cell, "columnNumber").and.returnValue(6);
        });

        it("should throw an error if the row or column offset is not an integer", () => {
            expect(() => cell.relativeCell()).toThrow();
            expect(() => cell.relativeCell("foo", 1)).toThrow();
            expect(() => cell.relativeCell(1, "foo")).toThrow();
            expect(() => cell.relativeCell(2.5, 1)).toThrow();
            expect(() => cell.relativeCell(1, 2.5)).toThrow();
        });

        it("should throw an error if the row or column absolute position is less than 0", () => {
            expect(() => cell.relativeCell(-100, -100)).toThrow();
        });

        it("should return a cell relative to this one", () => {
            expect(cell.relativeCell(3, 7)).toBe("CELL");
            expect(sheet.cell).toHaveBeenCalledWith(8, 13);
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

    describe("value", () => {
        it("should return the cell after setting the value", () => {
            expect(cell.value(5)).toBe(cell);
        });

        it("should store a number", () => {
            cell.value(57.8);
            expect(cellNode.toString()).toBe('<c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="C5"><v>57.8</v></c>');

            cell.value(-6);
            expect(cellNode.toString()).toBe('<c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="C5"><v>-6</v></c>');

            cell.value(0);
            expect(cellNode.toString()).toBe('<c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="C5"><v>0</v></c>');
        });

        it("should store a boolean", () => {
            cell.value(true);
            expect(cellNode.toString()).toBe('<c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="C5" t="b"><v>1</v></c>');

            cell.value(false);
            expect(cellNode.toString()).toBe('<c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="C5" t="b"><v>0</v></c>');
        });

        it("should store a string", () => {
            cell.value("some string");
            expect(cellNode.toString()).toBe('<c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="C5" t="inlineStr"><is><t>some string</t></is></c>');
        });

        it("should store a date", () => {
            const date = new Date('01 Jan 2016 00:00:00');
            cell.value(date);
            expect(cellNode.toString()).toBe('<c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="C5"><v>EXCEL_NUMBER</v></c>');
            expect(utils.dateToExcelNumber).toHaveBeenCalledWith(date);
        });

        it("should clear the cell if null or undefined", () => {
            cell.value(null);
            expect(cellNode.toString()).toBe('<c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="C5"/>');

            cell.value();
            expect(cellNode.toString()).toBe('<c xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" r="C5"/>');
        });

        it("should throw and error is a different value is set", () => {
            expect(() => cell.value({})).toThrow();
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
});
