"use strict";

const proxyquire = require("proxyquire").noCallThru();

describe("Column", () => {
    let Column, column, columnNode, sheet;

    beforeEach(() => {
        Column = proxyquire("../lib/Column", {});
        sheet = jasmine.createSpyObj('sheet', ['cell', 'name', 'workbook']);
        sheet.cell.and.returnValue('CELL');
        sheet.name.and.returnValue('NAME');
        sheet.workbook.and.returnValue('WORKBOOK');

        columnNode = {
            name: 'col',
            attributes: {
                min: 5,
                max: 5
            }
        };

        column = new Column(sheet, columnNode);
    });

    describe("address", () => {
        it("should return the address", () => {
            expect(column.address()).toBe('E:E');
            expect(column.address({ anchored: true })).toBe('$E:$E');
            expect(column.address({ includeSheetName: true })).toBe('NAME!E:E');
            expect(column.address({ includeSheetName: true, anchored: true })).toBe('NAME!$E:$E');
        });
    });

    describe("cell", () => {
        it("should return a cell", () => {
            expect(column.cell(7)).toBe('CELL');
            expect(sheet.cell).toHaveBeenCalledWith(7, 5);
        });
    });

    describe("columnName", () => {
        it("should return the column name", () => {
            expect(column.columnName()).toBe('E');
        });
    });

    describe("columnNumber", () => {
        it("should return the column number", () => {
            expect(column.columnNumber()).toBe(5);
        });
    });

    describe("hidden", () => {
        it("should get/set hidden", () => {
            expect(column.hidden()).toBe(false);

            column.hidden(true);
            expect(column.hidden()).toBe(true);
            expect(columnNode).toEqualJson({
                name: 'col',
                attributes: {
                    min: 5,
                    max: 5,
                    hidden: 1
                }
            });

            column.hidden(false);
            expect(column.hidden()).toBe(false);
            expect(columnNode).toEqualJson({
                name: 'col',
                attributes: {
                    min: 5,
                    max: 5
                }
            });
        });
    });

    describe("sheet", () => {
        it("should return the sheet", () => {
            expect(column.sheet()).toBe(sheet);
        });
    });

    describe("toObject", () => {
        it("should return the object representation", () => {
            expect(column.toObject()).toBe(columnNode);
        });
    });

    describe("width", () => {
        it("should get/set the width", () => {
            expect(column.width()).toBeUndefined();

            column.width(56.7);
            expect(column.width()).toBe(56.7);
            expect(columnNode).toEqualJson({
                name: 'col',
                attributes: {
                    min: 5,
                    max: 5,
                    customWidth: 1,
                    width: 56.7
                }
            });

            column.width(undefined);
            expect(column.width()).toBeUndefined();
            expect(columnNode).toEqualJson({
                name: 'col',
                attributes: {
                    min: 5,
                    max: 5
                }
            });
        });
    });

    describe("workbook", () => {
        it("should return the workbook", () => {
            expect(column.workbook()).toBe("WORKBOOK");
        });
    });
});
