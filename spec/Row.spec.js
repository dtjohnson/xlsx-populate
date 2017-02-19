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

        Row = proxyquire("../lib/Row", { './Cell': Cell });
        sheet = jasmine.createSpyObj('sheet', ['name', 'workbook', 'existingColumnStyleId']);
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
        beforeEach(() => {
            Cell.calls.reset();
        });

        it("should return an existing cell", () => {
            expect(row.cell(2)).toEqual(jasmine.any(Cell));
            expect(Cell).not.toHaveBeenCalled();
        });

        it("should return an existing cell", () => {
            expect(row.cell('B')).toEqual(jasmine.any(Cell));
            expect(Cell).not.toHaveBeenCalled();
        });

        it("should create a new cell as needed", () => {
            expect(row.cell(5)).toEqual(jasmine.any(Cell));
            expect(Cell).toHaveBeenCalledWith(row, {
                name: 'c', attributes: { r: "E7" }, children: []
            });
        });

        it("should create a new cell as needed", () => {
            expect(row.cell('C')).toEqual(jasmine.any(Cell));
            expect(Cell).toHaveBeenCalledWith(row, {
                name: 'c', attributes: { r: "C7" }, children: []
            });
        });

        it("should create a new cell with an existing column style id", () => {
            sheet.existingColumnStyleId.and.returnValue(5);
            expect(row.cell('C')).toEqual(jasmine.any(Cell));
            expect(Cell).toHaveBeenCalledWith(row, {
                name: 'c', attributes: { r: "C7", s: 5 }, children: []
            });
        });

        it("should create a new cell with an existing row style id", () => {
            rowNode.attributes.s = 3;
            expect(row.cell('C')).toEqual(jasmine.any(Cell));
            expect(Cell).toHaveBeenCalledWith(row, {
                name: 'c', attributes: { r: "C7", s: 3 }, children: []
            });
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

    describe("hidden", () => {
        it("should get/set hidden", () => {
            expect(row.hidden()).toBe(false);

            row.hidden(true);
            expect(row.hidden()).toBe(true);
            expect(rowNode.attributes).toEqualJson({
                r: 7,
                hidden: 1
            });

            row.hidden(false);
            expect(row.hidden()).toBe(false);
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

    describe("minUsedColumnNumber", () => {
        it("should return the min column number", () => {
            row._cells = [];
            row._cells[5] = row._cells[7] = {};
            expect(row.minUsedColumnNumber()).toBe(5);
        });
    });

    describe("maxUsedColumnNumber", () => {
        it("should return the max column number", () => {
            row._cells = [];
            row._cells[5] = row._cells[7] = {};
            expect(row.maxUsedColumnNumber()).toBe(7);
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

    describe("_init", () => {
        it("should store existing rows", () => {
            expect(row._cells).toEqual([undefined, undefined, jasmine.any(Cell)]);
            expect(Cell).toHaveBeenCalledWith(row, rowNode.children[0]);
        });
    });
});
