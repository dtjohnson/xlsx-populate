/* jshint jasmine: true */

"use strict";

var Cell = require('../lib/Cell'),
    etree = require('elementtree'),
    element = etree.Element,
    subelement = etree.SubElement;

describe("Cell", function () {
    var cell;
    var sheetMock = {
        getName: function () {
            return "Foo";
        }
    };

    beforeEach(function () {
        var c = element('c');
        c.attrib.t = "inlineStr";
        c.attrib.r = "C5";
        var isNode = subelement(c, "is");
        var tNode = subelement(isNode, "t");
        tNode.text = "Foo";
        cell = new Cell(sheetMock, 5, 3, c);
    });

    describe("getSheet", function () {
        it("should return the parent sheet object", function () {
            expect(cell.getSheet()).toBe(sheetMock);
        });
    });

    describe("getRow", function () {
        it("should return the row", function () {
            expect(cell.getRow()).toBe(5);
        });
    });

    describe("getColumn", function () {
        it("should return the column", function () {
            expect(cell.getColumn()).toBe(3);
        });
    });

    describe("getAddress", function () {
        it("should return the address", function () {
            expect(cell.getAddress()).toBe("C5");
        });
    });

    describe("getFullAddress", function () {
        it("should return the full address", function () {
            expect(cell.getFullAddress()).toBe("'Foo'!C5");
        });
    });

    describe("setValue", function () {
    });

    describe("setFormula", function () {
    });

    describe("_clearContents", function () {
        it("should clear the node contents", function () {
            expect(cell._cellNode.findall('*').length).toBe(1);
            expect(cell._cellNode.attrib.t).toBeTruthy();
            cell._clearContents();
            expect(cell._cellNode.findall('*').length).toBe(0);
            expect(cell._cellNode.attrib.t).toBeUndefined();
        });
    });
});
