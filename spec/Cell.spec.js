"use strict";

const proxyquire = require("proxyquire").noCallThru();

fdescribe("Cell", () => {
    let Cell, cell, cellNode, row, sheet;

    beforeEach(() => {
        Cell = proxyquire("../lib/Cell", {});
        sheet = jasmine.createSpyObj('sheet', ['createStyle', 'updateMaxSharedFormulaId', 'name', 'column', 'clearCellsUsingSharedFormula', 'cell']);
        sheet.name.and.returnValue("NAME");
        sheet.column.and.returnValue("COLUMN");
        row = jasmine.createSpyObj('row', ['sheet', 'workbook']);
        row.sheet.and.returnValue(sheet);
        row.workbook.and.returnValue("WORKBOOK");

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

    describe("formula", () => {
        it("should return undefined if formula not set", () => {
            expect(cell.formula()).toBeUndefined();
        });

        it("should return the formula if set", () => {
            cellNode.attributes = {};
            cellNode.children = [{
                name: 'f',
                attributes: {},
                children: ['FORMULA']
            }];

            expect(cell.formula()).toBe("FORMULA");
        });

        it("should return 'SHARED' if shared formula set", () => {
            cellNode.attributes = {};
            cellNode.children = [{
                name: 'f',
                attributes: { t: 'shared' },
                children: []
            }];

            expect(cell.formula()).toBe("SHARED");
        });

        it("should clear the formula", () => {
            cellNode.attributes.t = "b";
            cellNode.children = [{
                name: 'v',
                attributes: {},
                children: [1]
            }];

            expect(cell.formula(undefined)).toBe(cell);
            expect(cellNode).toEqualJson({
                name: 'c',
                attributes: { r: "C7" },
                children: []
            });
        });

        it("should set the formula and clear the value", () => {
            cellNode.attributes.t = "b";
            cellNode.children = [{
                name: 'v',
                attributes: {},
                children: [1]
            }];

            expect(cell.formula("FORMULA")).toBe(cell);
            expect(cellNode).toEqualJson({
                name: 'c',
                attributes: { r: "C7" },
                children: [{
                    name: 'f',
                    attributes: {},
                    children: ["FORMULA"]
                }]
            });
        });
    });

    describe("find", () => {
        beforeEach(() => {
            spyOn(cell, 'value');
        });

        it("should return true if substring found and false otherwise", () => {
            cell.value.and.returnValue("Foo bar baz");
            expect(cell.find('bar')).toBe(true);
            expect(cell.find('BAR')).toBe(true);
            expect(cell.find('goo')).toBe(false);
        });

        it("should return true if regex matches and false otherwise", () => {
            cell.value.and.returnValue("Foo bar baz");
            expect(cell.find(/\w{3}/)).toBe(true);
            expect(cell.find(/\w{4}/)).toBe(false);
            expect(cell.find(/Foo/)).toBe(true);
        });

        it("should replace all occurences of substring", () => {
            cell.value.and.returnValue("Foo bar baz foo");
            expect(cell.find('foo', 'XXX')).toBe(true);
            expect(cell.value).toHaveBeenCalledWith('XXX bar baz XXX');
            cell.value.calls.reset();

            cell.value.and.returnValue("Foo bar baz foo");
            expect(cell.find('foot', 'XXX')).toBe(false);
            expect(cell.value).not.toHaveBeenCalledWith(jasmine.any(String));
        });

        it("should replace regex matches", () => {
            cell.value.and.returnValue("Foo bar baz foo");
            expect(cell.find(/[a-z]{3}/, 'XXX')).toBe(true);
            expect(cell.value).toHaveBeenCalledWith('Foo XXX baz foo');
        });

        it("should replace regex matches", () => {
            cell.value.and.returnValue("Foo bar baz foo");
            const replacer = jasmine.createSpy('replacer').and.returnValue("REPLACEMENT");
            expect(cell.find(/(\w)(o{2})/g, replacer)).toBe(true);
            expect(cell.value).toHaveBeenCalledWith('REPLACEMENT bar baz REPLACEMENT');
            expect(replacer).toHaveBeenCalledWith('Foo', 'F', 'oo', 0, 'Foo bar baz foo');
            expect(replacer).toHaveBeenCalledWith('foo', 'f', 'oo', 12, 'Foo bar baz foo');
        });
    });

    xdescribe("groupWith", () => {
        it("should");
    });

    describe("tap", () => {
        it("should call the callback and return the cell", () => {
            const callback = jasmine.createSpy('callback').and.returnValue("RETURN");
            expect(cell.tap(callback)).toBe(cell);
            expect(callback).toHaveBeenCalledWith(cell);
        });
    });

    describe("thru", () => {
        it("should call the callback and return the callback return value", () => {
            const callback = jasmine.createSpy('callback').and.returnValue("RETURN");
            expect(cell.thru(callback)).toBe("RETURN");
            expect(callback).toHaveBeenCalledWith(cell);
        });
    });

    xdescribe("rangeTo", () => {
        it("should");
    });

    describe("relativeCell", () => {
        it("should call sheet.cell with the appropriate row/column", () => {
            sheet.cell.and.returnValue("CELL");

            expect(cell.relativeCell(0, 0)).toBe("CELL");
            expect(sheet.cell).toHaveBeenCalledWith(7, 3);
            sheet.cell.calls.reset();

            expect(cell.relativeCell(-2, -1)).toBe("CELL");
            expect(sheet.cell).toHaveBeenCalledWith(5, 2);
            sheet.cell.calls.reset();

            expect(cell.relativeCell(5, 2)).toBe("CELL");
            expect(sheet.cell).toHaveBeenCalledWith(12, 5);
            sheet.cell.calls.reset();
        });
    });

    describe("row", () => {
        it("should return the parent row", () => {
            expect(cell.row()).toBe(row);
        });
    });

    describe("rowNumber", () => {
        it("should return the row number", () => {
            expect(cell.rowNumber()).toBe(7);
        });
    });

    describe("sheet", () => {
        it("should return the parent sheet", () => {
            expect(cell.sheet()).toBe(sheet);
        });
    });

    describe("style", () => {

    });

    describe("value", () => {

    });

    describe("workbook", () => {
        it("should return the workbook from the row", () => {
            expect(cell.workbook()).toBe("WORKBOOK");
        });
    });

    describe("sharesFormula", () => {

    });

    describe("setSharedFormula", () => {

    });

    describe("toObject", () => {
        it("should return the node", () => {
            expect(cell.toObject()).toBe(cellNode);
        });
    });

    describe("_getSharedFormulaRefId", () => {
        it("should return -1 if not a formula", () => {
            expect(cell._getSharedFormulaRefId()).toBe(-1);
        });

        it("should return -1 if not a shared formula ref", () => {
            cellNode.children = [{
                name: 'f',
                attributes: {}
            }];
            expect(cell._getSharedFormulaRefId()).toBe(-1);
        });

        it("should return the shared formula ID", () => {
            cellNode.children = [{
                name: 'f',
                attributes: { ref: "A1:C3", si: 5 }
            }];
            expect(cell._getSharedFormulaRefId()).toBe(5);
        });
    });

    describe("_initNode", () => {

    });
});
