"use strict";

const proxyquire = require("proxyquire").noCallThru();

describe("Row", () => {
    let Row, Cell, row, rowNode, sheet;

    beforeEach(() => {
        let i = 1;
        Cell = jasmine.createSpy("Cell").and.callFake(function () {
            this.id = i++;
        });
        Cell.prototype.columnNumber = jasmine.createSpy("columnNumber").and.returnValue(2);
        Cell.prototype.toObject = jasmine.createSpy("toObject").and.callFake(function () {
            return this.id;
        });
        Cell.prototype.find = jasmine.createSpy('find');
        Cell.prototype.replace = jasmine.createSpy('replace');

        Row = proxyquire("../lib/Row", { './Cell': Cell });
        sheet = jasmine.createSpyObj('sheet', ['name', 'workbook']);
        sheet.name.and.returnValue('NAME');
        sheet.workbook.and.returnValue('WORKBOOK');

        rowNode = {
            name: 'row',
            attributes: {
                r: 7
            },
            children: [{
                name: 'c',
                attributes: {
                    r: "B7"
                }
            }]
        };

        row = new Row(sheet, rowNode);
    });

    describe("address", () => {
        it("should return the address", () => {
            expect(row.address()).toBe('7:7');
            expect(row.address({ anchored: true })).toBe('$7:$7');
            expect(row.address({ includeSheetName: true })).toBe('NAME!7:7');
            expect(row.address({ includeSheetName: true, anchored: true })).toBe('NAME!$7:$7');
        });
    });

    describe("cell", () => {
        it("should return an existing cell", () => {
            Cell.calls.reset();
            expect(row.cell(2)).toEqual(jasmine.any(Cell));
            expect(Cell).not.toHaveBeenCalled();
        });

        it("should return an existing cell", () => {
            Cell.calls.reset();
            expect(row.cell('B')).toEqual(jasmine.any(Cell));
            expect(Cell).not.toHaveBeenCalled();
        });

        it("should create a new cell as needed", () => {
            Cell.calls.reset();
            expect(row.cell(5)).toEqual(jasmine.any(Cell));
            expect(Cell).toHaveBeenCalledWith(row, {
                name: 'c', attributes: { r: "E7" }, children: []
            });
        });

        it("should create a new cell as needed", () => {
            Cell.calls.reset();
            expect(row.cell('C')).toEqual(jasmine.any(Cell));
            expect(Cell).toHaveBeenCalledWith(row, {
                name: 'c', attributes: { r: "C7" }, children: []
            });
        });
    });

    describe("find", () => {
        it("should return the matches", () => {
            Cell.prototype.find.and.returnValue(true);
            expect(row.find('foo')).toEqual([row.cell(2)]);
            expect(Cell.prototype.find).toHaveBeenCalledWith(/foo/gim, undefined);

            Cell.prototype.find.and.returnValue(false);
            expect(row.find('bar', 'baz')).toEqual([]);
            expect(Cell.prototype.find).toHaveBeenCalledWith(/bar/gim, 'baz');
        });
    });

    describe("height", () => {
        it("should get/set the height", () => {
            expect(row.height()).toBeUndefined();

            row.height(56.7);
            expect(row.height()).toBe(56.7);
            expect(rowNode.attributes).toEqualJson({
                r: 7,
                customHeight: 1,
                ht: 56.7
            });

            row.height(undefined);
            expect(row.height()).toBeUndefined();
            expect(rowNode.attributes).toEqualJson({
                r: 7
            });
        });
    });

    describe("rowNumber", () => {
        it("should return the row number", () => {
            expect(row.rowNumber()).toBe(7);
        });
    });

    describe("sheet", () => {
        it("should return the sheet", () => {
            expect(row.sheet()).toBe(sheet);
        });
    });

    describe("toObject", () => {
        it("should return the object representation with children in order", () => {
            row.cell(3);
            row.cell(1);
            const obj = row.toObject();
            expect(obj).toBe(rowNode);
            expect(obj.children).toEqualJson([3, 1, 2]);
        });
    });

    describe("workbook", () => {
        it("should return the workbook", () => {
            expect(row.workbook()).toBe("WORKBOOK");
        });
    });

    describe("_initNode", () => {
        it("should store existing rows", () => {
            expect(row._cells).toEqual([undefined, undefined, jasmine.any(Cell)]);
            expect(Cell).toHaveBeenCalledWith(row, rowNode.children[0]);
        });
    });
});
