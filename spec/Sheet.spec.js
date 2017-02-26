"use strict";

const proxyquire = require("proxyquire");

describe("Sheet", () => {
    let Sheet, Range, Row, Column, sheet, idNode, sheetNode, workbook;

    beforeEach(() => {
        let i = 0;
        workbook = jasmine.createSpyObj("workbook", ["scopedDefinedName"]);
        workbook.scopedDefinedName.and.returnValue("DEFINED NAME");

        Range = jasmine.createSpy("Range");
        Row = jasmine.createSpy("Row");
        Row.prototype.rowNumber = jasmine.createSpy().and.callFake(() => ++i);
        Row.prototype.find = jasmine.createSpy('find');
        Column = jasmine.createSpy("Column");

        Sheet = proxyquire("../lib/Sheet", {
            './Range': Range,
            './Row': Row,
            './Column': Column,
            '@noCallThru': true
        });

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
                xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
            },
            children: [
                { name: 'sheetFormatPr', attributes: {}, children: [] },
                { name: 'sheetData', attributes: {}, children: [] },
                { name: 'pageMargins', attributes: {}, children: [] }
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

    describe("definedName", () => {
        it("should return the defined name", () => {
            expect(sheet.definedName("FOO")).toBe("DEFINED NAME");
            expect(workbook.scopedDefinedName).toHaveBeenCalledWith("FOO", sheet);
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

    describe("range", () => {
        beforeEach(() => {
            spyOn(sheet, "cell").and.callFake((a, b) => [a, b]);
        });

        it("should create a range from a range address", () => {
            expect(sheet.range("A2:B3")).toEqual(jasmine.any(Range));
            expect(Range).toHaveBeenCalledWith([2, 1], [3, 2]);
        });

        it("should create a range from two cells or addresses", () => {
            const c1 = {}, c2 = {};
            expect(sheet.range(c1, c2)).toEqual(jasmine.any(Range));
            expect(Range).toHaveBeenCalledWith(c1, c2);

            expect(sheet.range("A1", c2)).toEqual(jasmine.any(Range));
            expect(Range).toHaveBeenCalledWith(['A1', undefined], c2);

            expect(sheet.range(c1, "C3")).toEqual(jasmine.any(Range));
            expect(Range).toHaveBeenCalledWith(c1, ["C3", undefined]);

            expect(sheet.range("A1", "C3")).toEqual(jasmine.any(Range));
            expect(Range).toHaveBeenCalledWith(["A1", undefined], ["C3", undefined]);
        });

        it("should create a range from row numbers and column names and numbers", () => {
            expect(sheet.range(1, 2, 3, 4)).toEqual(jasmine.any(Range));
            expect(Range).toHaveBeenCalledWith([1, 2], [3, 4]);

            expect(sheet.range(1, 'B', 3, 'D')).toEqual(jasmine.any(Range));
            expect(Range).toHaveBeenCalledWith([1, 'B'], [3, 'D']);
        });
    });

    describe("row", () => {
        it("should return an existing row", () => {
            const row = sheet._rows[3] = {};
            expect(sheet.row(3)).toBe(row);
        });

        it("should create a new row", () => {
            const row = sheet.row(5);
            expect(row).toEqual(jasmine.any(Row));
            expect(sheet._rows[5]).toBe(row);
            expect(Row).toHaveBeenCalledWith(sheet, {
                name: 'row',
                attributes: {
                    r: 5
                },
                children: []
            });
        });
    });

    describe("usedRange", () => {
        beforeEach(() => {
            spyOn(sheet, "range").and.returnValue("RANGE");
        });

        it("should return undefined", () => {
            sheet._rows = [];
            expect(sheet.usedRange()).toBeUndefined();

            sheet._rows = {
                minUsedColumnNumber: () => -1,
                maxUsedColumnNumber: () => 0
            };
            expect(sheet.usedRange()).toBeUndefined();
        });

        it("should return the used range", () => {
            sheet._rows = [
                undefined,
                undefined,
                undefined,
                {
                    minUsedColumnNumber: () => 3,
                    maxUsedColumnNumber: () => 5
                },
                undefined,
                undefined,
                {
                    minUsedColumnNumber: () => 2,
                    maxUsedColumnNumber: () => 4
                }
            ];

            expect(sheet.usedRange()).toBe("RANGE");
            expect(sheet.range).toHaveBeenCalledWith(3, 2, 6, 5);
        });
    });

    describe("workbook", () => {
        it("should get the workbook", () => {
            expect(sheet.workbook()).toBe(workbook);
        });
    });

    describe("areCellsMerged", () => {
        it("should return true/false if the cells are merged or not", () => {
            sheet._mergeCells = {
                ADDRESS1: {},
                ADDRESS2: {}
            };

            expect(sheet.areCellsMerged("ADDRESS1")).toBe(true);
            expect(sheet.areCellsMerged("ADDRESS2")).toBe(true);
            expect(sheet.areCellsMerged("ADDRESS3")).toBe(false);
        });
    });

    describe("clearCellsUsingSharedFormula", () => {
        it("should clear cells with matching shared formula", () => {
            sheet._rows = [
                undefined,
                {
                    clearCellsUsingSharedFormula: jasmine.createSpy("clearCellsUsingSharedFormula")
                },
                undefined,
                {
                    clearCellsUsingSharedFormula: jasmine.createSpy("clearCellsUsingSharedFormula")
                }
            ];

            sheet.clearCellsUsingSharedFormula(3);
            expect(sheet._rows[1].clearCellsUsingSharedFormula).toHaveBeenCalledWith(3);
            expect(sheet._rows[3].clearCellsUsingSharedFormula).toHaveBeenCalledWith(3);
        });
    });

    describe("existingColumnStyleId", () => {
        it("should return undefined if no existing column", () => {
            expect(sheet.existingColumnStyleId(3)).toBeUndefined();
        });

        it("should return the style ID from the column", () => {
            sheet._columns[5] = {
                styleId: () => "STYLE ID"
            };
            expect(sheet.existingColumnStyleId(5)).toBe("STYLE ID");
        });
    });

    describe("incrementMaxSharedFormulaId", () => {
        it("should increment the max shared formula ID", () => {
            sheet._maxSharedFormulaId = 8;
            expect(sheet.incrementMaxSharedFormulaId()).toBe(9);
            expect(sheet.incrementMaxSharedFormulaId()).toBe(10);
            expect(sheet.incrementMaxSharedFormulaId()).toBe(11);
        });
    });

    describe("mergeCells", () => {
        it("should add a mergeCell entry", () => {
            expect(sheet._mergeCells).toEqualJson({});
            sheet.mergeCells("ADDRESS");
            expect(sheet._mergeCells).toEqualJson({
                ADDRESS: {
                    name: 'mergeCell',
                    attributes: {
                        ref: "ADDRESS"
                    },
                    children: []
                }
            });
        });
    });

    describe("toObject", () => {
        it("should add the rows", () => {
            sheet._rows = [
                undefined,
                { toObject: () => "ROW1" },
                undefined,
                { toObject: () => "ROW2" }
            ];
            expect(sheet.toObject().children).toEqualJson([
                { name: 'sheetFormatPr', attributes: {}, children: [] },
                {
                    name: 'sheetData',
                    attributes: {},
                    children: ["ROW1", "ROW2"]
                },
                { name: 'pageMargins', attributes: {}, children: [] }
            ]);
        });

        it("should add the columns", () => {
            sheet._columns = [
                undefined,
                { toObject: () => "COLUMN1" },
                undefined,
                { toObject: () => "COLUMN2" }
            ];
            expect(sheet.toObject().children).toEqualJson([
                { name: 'sheetFormatPr', attributes: {}, children: [] },
                {
                    name: 'cols',
                    attributes: {},
                    children: ["COLUMN1", "COLUMN2"]
                },
                { name: 'sheetData', attributes: {}, children: [] },
                { name: 'pageMargins', attributes: {}, children: [] }
            ]);
        });

        it("should add the mergeCells", () => {
            sheet._mergeCells = {
                "A1:B2": "MERGE1",
                "C3:D4": "MERGE2"
            };

            expect(sheet.toObject().children).toEqualJson([
                { name: 'sheetFormatPr', attributes: {}, children: [] },
                { name: 'sheetData', attributes: {}, children: [] },
                {
                    name: 'mergeCells',
                    attributes: {},
                    children: ["MERGE1", "MERGE2"]
                },
                { name: 'pageMargins', attributes: {}, children: [] }
            ]);
        });
    });

    describe("unmergeCells", () => {
        it("should remove the mergeCell entry", () => {
            sheet._mergeCells = {
                ADDRESS1: {},
                ADDRESS2: {}
            };

            sheet.unmergeCells("ADDRESS3");
            expect(sheet._mergeCells).toEqualJson({
                ADDRESS1: {},
                ADDRESS2: {}
            });

            sheet.unmergeCells("ADDRESS1");
            expect(sheet._mergeCells).toEqualJson({
                ADDRESS2: {}
            });

            sheet.unmergeCells("ADDRESS2");
            expect(sheet._mergeCells).toEqualJson({});
        });
    });

    describe("updateMaxSharedFormulaId", () => {
        it("should update the max ID if greater", () => {
            sheet.updateMaxSharedFormulaId(5);
            expect(sheet._maxSharedFormulaId).toBe(5);

            sheet.updateMaxSharedFormulaId(3);
            expect(sheet._maxSharedFormulaId).toBe(5);

            sheet.updateMaxSharedFormulaId(undefined);
            expect(sheet._maxSharedFormulaId).toBe(5);

            sheet.updateMaxSharedFormulaId(7);
            expect(sheet._maxSharedFormulaId).toBe(7);
        });
    });

    describe("_init", () => {
        it("should parse the rows", () => {
            sheet._init({}, {}, {
                attributes: {},
                children: [
                    {
                        name: 'sheetData',
                        attributes: {},
                        children: ["ROW1", "ROW2", "ROW3"]
                    }
                ]
            });

            expect(sheet._rows).toEqual([
                undefined,
                jasmine.any(Row),
                jasmine.any(Row),
                jasmine.any(Row)
            ]);

            expect(Row).toHaveBeenCalledWith(sheet, "ROW1");
            expect(Row).toHaveBeenCalledWith(sheet, "ROW2");
            expect(Row).toHaveBeenCalledWith(sheet, "ROW3");
        });

        it("should parse the columns", () => {
            sheet._init({}, {}, {
                attributes: {},
                children: [
                    {
                        name: 'cols',
                        attributes: {},
                        children: [
                            { name: 'col', attributes: { min: 2, max: 3, foo: true } },
                            { name: 'col', attributes: { min: 5, max: 5, bar: true } }
                        ]
                    },
                    { name: "sheetData", attributes: {}, children: [] }
                ]
            });

            expect(sheet._columns).toEqual([
                undefined,
                undefined,
                jasmine.any(Column),
                jasmine.any(Column),
                undefined,
                jasmine.any(Column)
            ]);

            expect(Column).toHaveBeenCalledWith(sheet, { name: 'col', attributes: { min: 2, max: 2, foo: true } });
            expect(Column).toHaveBeenCalledWith(sheet, { name: 'col', attributes: { min: 3, max: 3, foo: true } });
            expect(Column).toHaveBeenCalledWith(sheet, { name: 'col', attributes: { min: 5, max: 5, bar: true } });
        });

        it("should parse the merged cells", () => {
            sheet._init({}, {}, {
                attributes: {},
                children: [
                    { name: "sheetData", attributes: {}, children: [] },
                    {
                        name: 'mergeCells',
                        attributes: {},
                        children: [
                            { name: 'mergeCell', attributes: { ref: "A1:B2", foo: true } },
                            { name: 'mergeCell', attributes: { ref: "C3:D4", bar: true } }
                        ]
                    }
                ]
            });

            expect(sheet._mergeCells).toEqualJson({
                "A1:B2": { name: 'mergeCell', attributes: { ref: "A1:B2", foo: true } },
                "C3:D4": { name: 'mergeCell', attributes: { ref: "C3:D4", bar: true } }
            });
        });
    });
});
