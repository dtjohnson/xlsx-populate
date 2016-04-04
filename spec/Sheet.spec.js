"use strict";

var proxyquire = require("proxyquire").noCallThru();
var xpath = require('xpath');
var DOMParser = require('xmldom').DOMParser;
var parser = new DOMParser();

describe("Sheet", function () {
    var Row, Sheet, workbook, sheetNode, sheetXML, sheet;

    beforeEach(function () {
        Row = jasmine.createSpy("Row");
        Sheet = proxyquire("../lib/Sheet", { './Row': Row });
        workbook = {};
        sheetNode = parser.parseFromString('<sheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="Sheet1" sheetId="1"/>').documentElement;
        sheetXML = parser.parseFromString('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"/></sheetData></worksheet>').documentElement;
        sheet = new Sheet(workbook, sheetNode, sheetXML);
    });

    describe("getWorkbook", function () {
        it("should return the workbook", function () {
            expect(sheet.getWorkbook()).toBe(workbook);
        });
    });

    describe("getName", function () {
        it("should return the sheet name", function () {
            expect(sheet.getName()).toBe("Sheet1");
        });
    });

    describe("getName", function () {
        it("should set the sheet name", function () {
            sheet.setName("some name");
            expect(sheetNode.toString()).toBe('<sheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="some name" sheetId="1"/>');
        });
    });

    describe("getRow", function () {
        it("should create a new row node if it doesn't exist", function () {
            sheet.getRow(3);
            expect(Row).toHaveBeenCalledWith(sheet, sheetXML.firstChild.lastChild);
            expect(sheetXML.toString()).toBe('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"/><row r="3"/></sheetData></worksheet>');
        });

        it("should use an existing row node if it does exist", function () {
            sheet.getRow(1);
            expect(Row).toHaveBeenCalledWith(sheet, sheetXML.firstChild.firstChild);
            expect(sheetXML.toString()).toBe('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"/></sheetData></worksheet>');
        });
    });

    describe("getCell", function () {
        var getCell;
        beforeEach(function () {
            getCell = jasmine.createSpy("getCell");
            sheet.getRow = jasmine.createSpy("getRow").and.returnValue({ getCell: getCell });
        });

        it("should call getRow and getCell with the given row and column", function () {
            sheet.getCell(5, 7);
            expect(sheet.getRow).toHaveBeenCalledWith(5);
            expect(getCell).toHaveBeenCalledWith(7);
        });

        it("should call getRow and getCell with the row and column corresponding to the given address", function () {
            sheet.getCell("H11");
            expect(sheet.getRow).toHaveBeenCalledWith(11);
            expect(getCell).toHaveBeenCalledWith(8);
        });
    });
});
