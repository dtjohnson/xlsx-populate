"use strict";

const proxyquire = require("proxyquire").noCallThru();

fdescribe("Cell", () => {
    let Cell, cell, cellNode, row, sheet;

    beforeEach(() => {
        Cell = proxyquire("../lib/Cell", {});
        sheet = jasmine.createSpyObj('sheet', ['createStyle', 'updateMaxSharedFormulaId', 'name', 'column', 'clearCellsUsingSharedFormula']);
        sheet.name.and.returnValue("NAME");
        sheet.column.and.returnValue("COLUMN");
        row = jasmine.createSpyObj('row', ['sheet', 'workbook']);
        row.sheet.and.returnValue(sheet);

        cellNode = {
            name: 'c',
            attributes: {
                r: "C7"
            }
        };

        cell = new Cell(row, cellNode);
    });

    describe("address", () => {
        it("should return the address", () => {
            expect(cell.address()).toBe('C7');
            expect(cell.address({ rowAnchored: true })).toBe('C$7');
            expect(cell.address({ columnAnchored: true })).toBe('$C7');
            expect(cell.address({ includeSheetName: true })).toBe('NAME!C7');
            expect(cell.address({ includeSheetName: true, rowAnchored: true, columnAnchored: true })).toBe('NAME!$C$7');
        });
    });

    describe("column", () => {
        it("should return the parent column", () => {
            expect(cell.column()).toBe("COLUMN");
            expect(sheet.column).toHaveBeenCalledWith(3);
        });
    });

    describe("clear", () => {
        it("should clear the cell contents", () => {
            cellNode.attributes.t = "TYPE";
            cellNode.children = ["CHILDREN"];
            expect(cell.clear()).toBe(cell);
            expect(cellNode.attributes.t).toBeUndefined();
            expect(cellNode.children).toEqual([]);
            expect(sheet.clearCellsUsingSharedFormula).not.toHaveBeenCalled();
        });

        it("should clear the cell with shared ref", () => {
            spyOn(cell, '_getSharedFormulaRefId').and.returnValue(4);
            cellNode.attributes.t = "TYPE";
            cellNode.children = ["CHILDREN"];
            expect(cell.clear()).toBe(cell);
            expect(cellNode.attributes.t).toBeUndefined();
            expect(cellNode.children).toEqual([]);
            expect(sheet.clearCellsUsingSharedFormula).toHaveBeenCalledWith(4);
        });
    });

    describe("columnName", () => {
        it("should return the column name", () => {
            expect(cell.columnName()).toBe("C");
        });
    });

    describe("columnNumber", () => {
        it("should return the column number", () => {
            expect(cell.columnNumber()).toBe(3);
        });
    });




    describe("sheet", () => {
        it("should return the parent sheet", () => {
            expect(cell.sheet()).toBe(sheet);
        });
    });
});
