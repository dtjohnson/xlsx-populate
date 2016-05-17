"use strict";

var utils = require("../lib/utils");
var Workbook = require("../lib/Workbook");

var proxyquire = require("proxyquire").noCallThru();
var DOMParser = require("xmldom").DOMParser;
var parser = new DOMParser();

var findChildNodes = function (node, tag) {
    var result = [];
    for (var childNodeIndex in node.childNodes) {
        if (node.childNodes.hasOwnProperty(childNodeIndex)) {
            var childNode = node.childNodes[childNodeIndex];
            if (childNode.tagName === tag) {
                result.push(childNode);
            }
        }
    }
    return result;
};

var getNodeText = function (node) {
    if (node.childNodes.length === 1) {
        return node.childNodes[0].nodeValue;
    }
    return null;
};

var getNodeAttribute = function (node, attribute) {
    return node.getAttribute(attribute);
};

describe("Sheet", function () {
    var Row, Sheet, workbook, sheetNode, sheetXML, sheet;

    beforeEach(function () {
        Row = jasmine.createSpy("Row");
        Sheet = proxyquire("../lib/Sheet", { "./Row": Row });
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

        it("should create rows in order", function () {
            sheet.getRow(3);
            sheet.getRow(2);
            expect(sheetXML.toString()).toBe('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"/><row r="2"/><row r="3"/></sheetData></worksheet>');
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

        it("should call getRow and getCell with the row and column corresponding to a lowercase address", function () {
            sheet.getCell("g9");
            expect(sheet.getRow).toHaveBeenCalledWith(9);
            expect(getCell).toHaveBeenCalledWith(7);
        });
    });

    xdescribe("getCell (advanced test)", function () {
        var workbook, sheet;
        beforeEach(function () {
            workbook = Workbook.fromBlankSync();
            sheet = workbook.getSheet(0);
        });

        describe("deterministic access", function () {
            it("correctly maps to the same cell", function () {
                var upperCaseCell = sheet.getCell("A1");
                var lowerCaseCell = sheet.getCell("a1");
                var rowAndColumnCell = sheet.getCell(1, 1);
                expect(upperCaseCell.getFullAddress()).toBe(lowerCaseCell.getFullAddress());
                expect(upperCaseCell.getFullAddress()).toBe(rowAndColumnCell.getFullAddress());
                var num = Math.random();
                upperCaseCell.setValue(num);
                var upperCaseVNode = findChildNodes(upperCaseCell._cellNode, "v")[0];
                var lowerCaseVNode = findChildNodes(lowerCaseCell._cellNode, "v")[0];
                var rowAndColumnVNode = findChildNodes(rowAndColumnCell._cellNode, "v")[0];
                expect(upperCaseVNode).not.toBeNull("A1 value node should not be null");
                expect(lowerCaseVNode).not.toBeNull("a1 value node should not be null");
                expect(rowAndColumnVNode).not.toBeNull("1,1 value node should not be null");
                expect(getNodeText(upperCaseVNode)).toBe(getNodeText(lowerCaseVNode));
                expect(getNodeText(upperCaseVNode)).toBe(getNodeText(rowAndColumnVNode));
                expect(parseFloat(getNodeText(upperCaseVNode))).toBe(num, "A1 value set must match value generated");
                expect(parseFloat(getNodeText(lowerCaseVNode))).toBe(num, "a1 value set must match value generated");
                expect(parseFloat(getNodeText(rowAndColumnVNode))).toBe(num, "1,1 value set must match value generated");
            });
        });

        describe("stochastic access", function () {
            var MAX_ROW = 100;
            var MAX_COLUMN = 100;
            var MAX_EDIT = 1000;

            beforeEach(function () {
                // Make random edits to the sheet
                for (var i = 0; i < MAX_EDIT; i++) {
                    var rowNumber = 1 + Math.floor(MAX_ROW * Math.random());
                    var columnNumber = 1 + Math.floor(MAX_COLUMN * Math.random());
                    var cell = sheet.getCell(rowNumber, columnNumber);
                    cell.setValue(Math.random());
                }
            });

            it("is stored in order", function () {
                // Reload workbook and sheet
                workbook = new Workbook(workbook.output());
                sheet = workbook.getSheet(0);

                // Check order
                var lastRowNumber = 0;
                var sheetDataNode = findChildNodes(sheet._sheetXML, "sheetData")[0];
                var rowNodes = findChildNodes(sheetDataNode, "row");
                expect(rowNodes.length).toBeGreaterThan(0, "Rows must exist after workbook reload");
                rowNodes.forEach(function (rowNode) {
                    var rowNumber = parseInt(getNodeAttribute(rowNode, "r"));
                    expect(isNaN(rowNumber)).toBe(false);
                    expect(rowNumber).toBeGreaterThan(lastRowNumber);
                    lastRowNumber = rowNumber;
                    var lastColumnNumber = 0;
                    var cNodes = findChildNodes(rowNode, "c");
                    cNodes.forEach(function (cNode) {
                        var address = getNodeAttribute(cNode, "r");
                        expect(address).toBeDefined();
                        var ref = utils.addressToRowAndColumn(address);
                        expect(ref.row).toBe(rowNumber);
                        expect(ref.column).toBeGreaterThan(lastColumnNumber);
                        lastColumnNumber = ref.column;
                    });
                });
            });

            it("does not contain duplicates", function () {
                var addressCounter = {};
                var sheetDataNode = findChildNodes(sheet._sheetXML, "sheetData")[0];
                var rowNodes = findChildNodes(sheetDataNode, "row");
                expect(rowNodes.length).toBeGreaterThan(0, "Rows must exist");
                rowNodes.forEach(function (rowNode) {
                    var cNodes = findChildNodes(rowNode, "c");
                    cNodes.forEach(function (cNode) {
                        var address = getNodeAttribute(cNode, "r");
                        expect(address).not.toBeNull();
                        expect(address).toBeDefined();
                        if (address in addressCounter === false) {
                            addressCounter[address] = 0;
                        }
                        addressCounter[address]++;
                    });
                });
                for (var address in addressCounter) {
                    if (addressCounter.hasOwnProperty(address)) {
                        expect(addressCounter[address]).toBeLessThan(2);
                    }
                }
            });
        });
    });
});
