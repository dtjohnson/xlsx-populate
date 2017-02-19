"use strict";

const proxyquire = require("proxyquire").noCallThru();

fdescribe("Sheet", () => {
    let Sheet, Range, Row, Column, sheet, idNode, sheetNode, workbook;

    beforeEach(() => {
        workbook = "WORKBOOK";

        Range = jasmine.createSpy("Range");
        Row = jasmine.createSpy("Row");
        Row.prototype.find = jasmine.createSpy('find');
        Column = jasmine.createSpy("Column");

        Sheet = proxyquire("../lib/Sheet", { './Range': Range, './Row': Row, './Column': Column });

        idNode = {
            name: 'sheet',
            attributes: {
                name: 'SHEET NAME',
                sheetId: '1',
                'r:id': 'rId1'
            },
            children: []
        };

        sheetNode = {
            name: 'worksheet',
            attributes: {
                xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
                'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
                'xmlns:mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
                'mc:Ignorable': 'x14ac',
                'xmlns:x14ac': 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac'
            },
            children: [
                { name: 'dimension', attributes: [Object], children: [] },
                { name: 'sheetViews', attributes: {}, children: [Object] },
                { name: 'sheetFormatPr', attributes: [Object], children: [] },
                { name: 'sheetData', attributes: {}, children: [] },
                { name: 'pageMargins', attributes: [Object], children: [] }
            ]
        };

        sheet = new Sheet(workbook, idNode, sheetNode);
    });

    describe("cell", () => {
        let cell;
        beforeEach(() => {
            cell = jasmine.createSpy("cell").and.returnValue("CELL");
            spyOn(sheet, "row").and.returnValue({ cell });
        });

        it("should create a cell from an address", () => {
            expect(sheet.cell("$B6")).toBe("CELL");
            expect(sheet.row).toHaveBeenCalledWith(6);
            expect(cell).toHaveBeenCalledWith(2);
        });

        it("should create a cell from a row/column", () => {
            expect(sheet.cell(5, 7)).toBe("CELL");
            expect(sheet.row).toHaveBeenCalledWith(5);
            expect(cell).toHaveBeenCalledWith(7);
        });
    });

    describe("column", () => {
        it("should get an existing column", () => {
            const column = sheet._columns[3] = {};
            expect(sheet.column(3)).toBe(column);
            expect(sheet.column('C')).toBe(column);
        });

        it("should create a new column", () => {
            const column = sheet.column('E');
            expect(column).toEqual(jasmine.any(Column));
            expect(sheet._columns[5]).toBe(column);
            expect(Column).toHaveBeenCalledWith(sheet, {
                name: 'col',
                attributes: {
                    min: 5,
                    max: 5
                }
            });
        });
    });

    describe("find", () => {
        it("should return the matches", () => {
            sheet.row(1);
            sheet.row(2);
            sheet.row(3);

            Row.prototype.find.and.returnValue(["A", "B"]);
            expect(sheet.find('foo')).toEqual(["A", "B", "A", "B", "A", "B"]);
            expect(Row.prototype.find).toHaveBeenCalledWith(/foo/gim, undefined);

            Row.prototype.find.and.returnValue('C');
            expect(sheet.find('bar', 'baz')).toEqual(['C', 'C', 'C']);
            expect(Row.prototype.find).toHaveBeenCalledWith(/bar/gim, 'baz');
        });
    });

    describe("name", () => {
        it("should return the sheet name", () => {
            expect(sheet.name()).toBe("SHEET NAME");
        });
    });
});
