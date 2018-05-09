"use strict";

const proxyquire = require("proxyquire");

describe("Range", () => {
    let Range, range, startCell, endCell, sheet, style;

    beforeEach(() => {
        Range = proxyquire("../../lib/Range", {
            '@noCallThru': true
        });

        const Style = class {}
        if (!Style.name) Style.name = "Style";
        Style.prototype.style = jasmine.createSpy("Style.style").and.callFake(name => `STYLE:${name}`);
        style = new Style();

        sheet = jasmine.createSpyObj('sheet', ['name', 'workbook', 'cell', 'merged', 'incrementMaxSharedFormulaId', 'dataValidation', 'autoFilter']);
        sheet.name.and.returnValue('NAME');
        sheet.cell.and.callFake((row, column) => `CELL[${row}, ${column}]`);
        sheet.workbook.and.returnValue('WORKBOOK');
        sheet.dataValidation.and.returnValue('DATAVALIDATION');

        startCell = jasmine.createSpyObj("startCell", ["rowNumber", "columnNumber", "columnName", "sheet", "value"]);
        startCell.columnName.and.returnValue("B");
        startCell.columnNumber.and.returnValue(2);
        startCell.rowNumber.and.returnValue(3);
        startCell.sheet.and.returnValue(sheet);

        endCell = jasmine.createSpyObj("endCell", ["rowNumber", "columnNumber", "columnName", "sheet", "value"]);
        endCell.columnName.and.returnValue("C");
        endCell.columnNumber.and.returnValue(3);
        endCell.rowNumber.and.returnValue(5);
        endCell.sheet.and.returnValue(sheet);

        range = new Range(startCell, endCell);
    });
    describe("address", () => {
        it("should return the address", () => {
            expect(range.address()).toBe('B3:C5');
            expect(range.address({ startRowAnchored: true })).toBe('B$3:C5');
            expect(range.address({ startColumnAnchored: true })).toBe('$B3:C5');
            expect(range.address({ endRowAnchored: true })).toBe('B3:C$5');
            expect(range.address({ endColumnAnchored: true })).toBe('B3:$C5');
            expect(range.address({ includeSheetName: true })).toBe("'NAME'!B3:C5");
            expect(range.address({
                includeSheetName: true,
                startRowAnchored: true,
                startColumnAnchored: true,
                endRowAnchored: true,
                endColumnAnchored: true
            })).toBe("'NAME'!$B$3:$C$5");
            expect(range.address({ anchored: true })).toBe("$B$3:$C$5");
        });
    });

    describe("cell", () => {
        it("should get the cell relative to the top/left corner", () => {
            expect(range.cell(0, 0)).toBe("CELL[3, 2]");
            expect(sheet.cell).toHaveBeenCalledWith(3, 2);

            expect(range.cell(-2, 3)).toBe("CELL[1, 5]");
            expect(sheet.cell).toHaveBeenCalledWith(1, 5);

            expect(range.cell(4, -1)).toBe("CELL[7, 1]");
            expect(sheet.cell).toHaveBeenCalledWith(7, 1);
        });
    });

    describe("cells", () => {
        it("should get the cells", () => {
            expect(range.cells()).toEqualJson([
                ["CELL[3, 2]", "CELL[3, 3]"],
                ["CELL[4, 2]", "CELL[4, 3]"],
                ["CELL[5, 2]", "CELL[5, 3]"]
            ]);
        });
    });

    describe("autoFilter", () => {
        it("should mark the range as having an automatic filter", () => {
            expect(range.autoFilter()).toBe(range);
        });
    });

    describe("clear", () => {
        it("should clear the cell", () => {
            spyOn(range, "value").and.returnValue("RETURN");
            expect(range.clear()).toBe("RETURN");
            expect(range.value).toHaveBeenCalledWith(undefined);
        });
    });

    describe("endCell", () => {
        it("should return the end cell", () => {
            expect(range.endCell()).toBe(endCell);
        });
    });

    describe("forEach", () => {
        it("should call the callback for each cell", () => {
            const callback = jasmine.createSpy("callback");
            expect(range.forEach(callback)).toBe(range);
            expect(callback.calls.argsFor(0)).toEqualJson(["CELL[3, 2]", 0, 0, range]);
            expect(callback.calls.argsFor(1)).toEqualJson(["CELL[3, 3]", 0, 1, range]);
            expect(callback.calls.argsFor(2)).toEqualJson(["CELL[4, 2]", 1, 0, range]);
            expect(callback.calls.argsFor(3)).toEqualJson(["CELL[4, 3]", 1, 1, range]);
            expect(callback.calls.argsFor(4)).toEqualJson(["CELL[5, 2]", 2, 0, range]);
            expect(callback.calls.argsFor(5)).toEqualJson(["CELL[5, 3]", 2, 1, range]);
        });
    });

    describe("formula", () => {
        it("should return the top-left cell shared ref formula", () => {
            spyOn(range, "startCell").and.returnValue({
                getSharedRefFormula: jasmine.createSpy("getSharedRefFormula").and.returnValue("RETURN")
            });

            expect(range.formula()).toBe("RETURN");
        });

        it("should set the shared formula", () => {
            sheet.incrementMaxSharedFormulaId.and.returnValue(8);
            const cells = [];
            sheet.cell.and.callFake((rowNumber, columnNumber) => {
                return cells[`${rowNumber}, ${columnNumber}`] = {
                    setSharedFormula: jasmine.createSpy("setSharedFormula")
                };
            });

            expect(range.formula("FORMULA")).toBe(range);
            expect(cells["3, 2"].setSharedFormula).toHaveBeenCalledWith(8, "FORMULA", "B3:C5");
            expect(cells["3, 3"].setSharedFormula).toHaveBeenCalledWith(8);
            expect(cells["4, 2"].setSharedFormula).toHaveBeenCalledWith(8);
            expect(cells["4, 3"].setSharedFormula).toHaveBeenCalledWith(8);
            expect(cells["5, 2"].setSharedFormula).toHaveBeenCalledWith(8);
            expect(cells["5, 3"].setSharedFormula).toHaveBeenCalledWith(8);
        });
    });

    describe("map", () => {
        it("should call the callback for each cell and return the values", () => {
            const callback = jasmine.createSpy("callback").and.callFake((cell, ri, ci) => `RETURN[${ri}, ${ci}]`);
            expect(range.map(callback)).toEqualJson([
                ["RETURN[0, 0]", "RETURN[0, 1]"],
                ["RETURN[1, 0]", "RETURN[1, 1]"],
                ["RETURN[2, 0]", "RETURN[2, 1]"]
            ]);
            expect(callback.calls.argsFor(0)).toEqualJson(["CELL[3, 2]", 0, 0, range]);
            expect(callback.calls.argsFor(1)).toEqualJson(["CELL[3, 3]", 0, 1, range]);
            expect(callback.calls.argsFor(2)).toEqualJson(["CELL[4, 2]", 1, 0, range]);
            expect(callback.calls.argsFor(3)).toEqualJson(["CELL[4, 3]", 1, 1, range]);
            expect(callback.calls.argsFor(4)).toEqualJson(["CELL[5, 2]", 2, 0, range]);
            expect(callback.calls.argsFor(5)).toEqualJson(["CELL[5, 3]", 2, 1, range]);
        });
    });

    describe("merged", () => {
        it("should get merged", () => {
            sheet.merged.and.returnValue("RETURN");
            expect(range.merged()).toBe("RETURN");
            expect(sheet.merged).toHaveBeenCalledWith("B3:C5");
        });

        it("should merge the cells", () => {
            expect(range.merged(true)).toBe(range);
            expect(sheet.merged).toHaveBeenCalledWith("B3:C5", true);
        });

        it("should unmerge the cells", () => {
            expect(range.merged(false)).toBe(range);
            expect(sheet.merged).toHaveBeenCalledWith("B3:C5", false);
        });
    });

    describe("reduce", () => {
        it("should call the callback for each cell and return the aggregate value", () => {
            const callback = jasmine.createSpy("callback").and.callFake((accumulator, cell, ri, ci) => `${accumulator} RETURN[${ri}, ${ci}]`);
            expect(range.reduce(callback, "INITIAL")).toBe("INITIAL RETURN[0, 0] RETURN[0, 1] RETURN[1, 0] RETURN[1, 1] RETURN[2, 0] RETURN[2, 1]");
            expect(callback.calls.argsFor(0)).toEqualJson(["INITIAL", "CELL[3, 2]", 0, 0, range]);
            expect(callback.calls.argsFor(1)).toEqualJson(["INITIAL RETURN[0, 0]", "CELL[3, 3]", 0, 1, range]);
            expect(callback.calls.argsFor(2)).toEqualJson(["INITIAL RETURN[0, 0] RETURN[0, 1]", "CELL[4, 2]", 1, 0, range]);
            expect(callback.calls.argsFor(3)).toEqualJson(["INITIAL RETURN[0, 0] RETURN[0, 1] RETURN[1, 0]", "CELL[4, 3]", 1, 1, range]);
            expect(callback.calls.argsFor(4)).toEqualJson(["INITIAL RETURN[0, 0] RETURN[0, 1] RETURN[1, 0] RETURN[1, 1]", "CELL[5, 2]", 2, 0, range]);
            expect(callback.calls.argsFor(5)).toEqualJson(["INITIAL RETURN[0, 0] RETURN[0, 1] RETURN[1, 0] RETURN[1, 1] RETURN[2, 0]", "CELL[5, 3]", 2, 1, range]);
        });
    });

    describe("sheet", () => {
        it("should return the sheet", () => {
            expect(range.sheet()).toBe(sheet);
        });
    });

    describe("startCell", () => {
        it("should return the end cell", () => {
            expect(range.endCell()).toBe(endCell);
        });
    });

    describe("style", () => {
        let cell;
        beforeEach(() => {
            cell = { style: style.style };
            sheet.cell.and.returnValue(cell);
        });

        it("should get a single style value", () => {
            expect(range.style("foo")).toEqualJson([
                ["STYLE:foo", "STYLE:foo"],
                ["STYLE:foo", "STYLE:foo"],
                ["STYLE:foo", "STYLE:foo"]
            ]);
            expect(cell.style).toHaveBeenCalledWith("foo");
        });

        it("should get multiple style values", () => {
            expect(range.style(["foo", "bar"])).toEqualJson({
                foo: [
                    ["STYLE:foo", "STYLE:foo"],
                    ["STYLE:foo", "STYLE:foo"],
                    ["STYLE:foo", "STYLE:foo"]
                ],
                bar: [
                    ["STYLE:bar", "STYLE:bar"],
                    ["STYLE:bar", "STYLE:bar"],
                    ["STYLE:bar", "STYLE:bar"]
                ]
            });
            expect(cell.style).toHaveBeenCalledWith("foo");
        });

        it("should set a style from the callback", () => {
            let i = 0;
            const callback = jasmine.createSpy("callback").and.callFake(() => i++);
            expect(range.style("foo", callback)).toBe(range);
            expect(cell.style).toHaveBeenCalledWith("foo", 0);
            expect(cell.style).toHaveBeenCalledWith("foo", 1);
            expect(cell.style).toHaveBeenCalledWith("foo", 2);
            expect(cell.style).toHaveBeenCalledWith("foo", 3);
            expect(cell.style).toHaveBeenCalledWith("foo", 4);
            expect(cell.style).toHaveBeenCalledWith("foo", 5);
            expect(callback).toHaveBeenCalledWith(cell, 0, 0, range);
            expect(callback).toHaveBeenCalledWith(cell, 0, 1, range);
            expect(callback).toHaveBeenCalledWith(cell, 1, 0, range);
            expect(callback).toHaveBeenCalledWith(cell, 1, 1, range);
            expect(callback).toHaveBeenCalledWith(cell, 2, 0, range);
            expect(callback).toHaveBeenCalledWith(cell, 2, 1, range);
        });

        it("should set a style from an array", () => {
            expect(range.style("foo", [
                [0, 1],
                [2, 3],
                [4, 5]
            ])).toBe(range);
            expect(cell.style).toHaveBeenCalledWith("foo", 0);
            expect(cell.style).toHaveBeenCalledWith("foo", 1);
            expect(cell.style).toHaveBeenCalledWith("foo", 2);
            expect(cell.style).toHaveBeenCalledWith("foo", 3);
            expect(cell.style).toHaveBeenCalledWith("foo", 4);
            expect(cell.style).toHaveBeenCalledWith("foo", 5);
        });

        it("should set a style from a single value", () => {
            expect(range.style("foo", "bar")).toBe(range);
            expect(cell.style).toHaveBeenCalledWith("foo", 'bar');
            expect(cell.style).toHaveBeenCalledWith("foo", 'bar');
            expect(cell.style).toHaveBeenCalledWith("foo", 'bar');
            expect(cell.style).toHaveBeenCalledWith("foo", 'bar');
            expect(cell.style).toHaveBeenCalledWith("foo", 'bar');
            expect(cell.style).toHaveBeenCalledWith("foo", 'bar');
        });

        it("should assign a style when asked", () => {
            expect(range.style(style)).toBe(range);
            expect(range._style).toBe(style);
            expect(cell.style).toHaveBeenCalledWith(style);
            expect(cell.style).toHaveBeenCalledWith(style);
            expect(cell.style).toHaveBeenCalledWith(style);
            expect(cell.style).toHaveBeenCalledWith(style);
            expect(cell.style).toHaveBeenCalledWith(style);
            expect(cell.style).toHaveBeenCalledWith(style);
        });

        it("should set styles from an object", () => {
            let i = 0;
            expect(range.style({
                foo: "FOO",
                bar: [["BAR0", "BAR1"], ["BAR2", "BAR3"], ["BAR4", "BAR5"]],
                baz: () => `BAZ${i++}`
            })).toBe(range);
            expect(cell.style).toHaveBeenCalledWith("foo", 'FOO');
            expect(cell.style).toHaveBeenCalledWith("foo", 'FOO');
            expect(cell.style).toHaveBeenCalledWith("foo", 'FOO');
            expect(cell.style).toHaveBeenCalledWith("foo", 'FOO');
            expect(cell.style).toHaveBeenCalledWith("foo", 'FOO');
            expect(cell.style).toHaveBeenCalledWith("foo", 'FOO');
            expect(cell.style).toHaveBeenCalledWith("bar", 'BAR0');
            expect(cell.style).toHaveBeenCalledWith("bar", 'BAR1');
            expect(cell.style).toHaveBeenCalledWith("bar", 'BAR2');
            expect(cell.style).toHaveBeenCalledWith("bar", 'BAR3');
            expect(cell.style).toHaveBeenCalledWith("bar", 'BAR4');
            expect(cell.style).toHaveBeenCalledWith("bar", 'BAR5');
            expect(cell.style).toHaveBeenCalledWith("baz", 'BAZ0');
            expect(cell.style).toHaveBeenCalledWith("baz", 'BAZ1');
            expect(cell.style).toHaveBeenCalledWith("baz", 'BAZ2');
            expect(cell.style).toHaveBeenCalledWith("baz", 'BAZ3');
            expect(cell.style).toHaveBeenCalledWith("baz", 'BAZ4');
            expect(cell.style).toHaveBeenCalledWith("baz", 'BAZ5');
        });
    });

    describe('dataValidation', () => {
        it('should return the range', () => {
            expect(range.dataValidation('testing, testing2')).toBe(range);
            expect(sheet.dataValidation).toHaveBeenCalledWith('B3:C5', 'testing, testing2');
        });

        it('should return the range', () => {
            expect(range.dataValidation({ type: 'list',
                allowBlank: false,
                showInputMessage: false,
                prompt: '',
                promptTitle: '',
                showErrorMessage: false,
                error: '',
                errorTitle: '',
                operator: '',
                formula1: 'test1, test2, test3',
                formula2: ''
            })).toBe(range);

            expect(sheet.dataValidation).toHaveBeenCalledWith('B3:C5', { type: 'list',
                allowBlank: false,
                showInputMessage: false,
                prompt: '',
                promptTitle: '',
                showErrorMessage: false,
                error: '',
                errorTitle: '',
                operator: '',
                formula1: 'test1, test2, test3',
                formula2: ''
            });
        })

        it("should get the dataValidation from the range", () => {
            expect(range.dataValidation()).toBe("DATAVALIDATION");
            expect(sheet.dataValidation).toHaveBeenCalledWith("B3:C5");
        });

    });

    describe("tap", () => {
        it("should call the callback and return the range", () => {
            const callback = jasmine.createSpy('callback').and.returnValue("RETURN");
            expect(range.tap(callback)).toBe(range);
            expect(callback).toHaveBeenCalledWith(range);
        });
    });

    describe("thru", () => {
        it("should call the callback and return the callback return value", () => {
            const callback = jasmine.createSpy('callback').and.returnValue("RETURN");
            expect(range.thru(callback)).toBe("RETURN");
            expect(callback).toHaveBeenCalledWith(range);
        });
    });

    describe("values", () => {
        let cell;
        beforeEach(() => {
            cell = { value: jasmine.createSpy("value").and.returnValue("VALUE") };
            sheet.cell.and.returnValue(cell);
        });

        it("should get the value", () => {
            expect(range.value()).toEqualJson([
                ["VALUE", "VALUE"],
                ["VALUE", "VALUE"],
                ["VALUE", "VALUE"]
            ]);
            expect(cell.value).toHaveBeenCalledWith();
        });

        it("should set the value from the callback", () => {
            let i = 0;
            const callback = jasmine.createSpy("callback").and.callFake(() => i++);
            expect(range.value(callback)).toBe(range);
            expect(cell.value).toHaveBeenCalledWith(0);
            expect(cell.value).toHaveBeenCalledWith(1);
            expect(cell.value).toHaveBeenCalledWith(2);
            expect(cell.value).toHaveBeenCalledWith(3);
            expect(cell.value).toHaveBeenCalledWith(4);
            expect(cell.value).toHaveBeenCalledWith(5);
            expect(callback).toHaveBeenCalledWith(cell, 0, 0, range);
            expect(callback).toHaveBeenCalledWith(cell, 0, 1, range);
            expect(callback).toHaveBeenCalledWith(cell, 1, 0, range);
            expect(callback).toHaveBeenCalledWith(cell, 1, 1, range);
            expect(callback).toHaveBeenCalledWith(cell, 2, 0, range);
            expect(callback).toHaveBeenCalledWith(cell, 2, 1, range);
        });

        it("should set values from an array", () => {
            expect(range.value([
                [0, 1],
                [2, 3],
                [4, 5]
            ])).toBe(range);
            expect(cell.value).toHaveBeenCalledWith(0);
            expect(cell.value).toHaveBeenCalledWith(1);
            expect(cell.value).toHaveBeenCalledWith(2);
            expect(cell.value).toHaveBeenCalledWith(3);
            expect(cell.value).toHaveBeenCalledWith(4);
            expect(cell.value).toHaveBeenCalledWith(5);
        });

        it("should set a single value", () => {
            expect(range.value("foo")).toBe(range);
            expect(cell.value).toHaveBeenCalledWith("foo");
            expect(cell.value).toHaveBeenCalledWith("foo");
            expect(cell.value).toHaveBeenCalledWith("foo");
            expect(cell.value).toHaveBeenCalledWith("foo");
            expect(cell.value).toHaveBeenCalledWith("foo");
            expect(cell.value).toHaveBeenCalledWith("foo");
        });
    });

    describe("workbook", () => {
        it("should return the workbook", () => {
            expect(range.workbook()).toBe("WORKBOOK");
        });
    });

    describe("_findRangeExtent", () => {
        it("should set the min/max row/column", () => {
            range._startCell = startCell;
            range._endCell = endCell;
            expect(range._minRowNumber).toBe(3);
            expect(range._maxRowNumber).toBe(5);
            expect(range._minColumnNumber).toBe(2);
            expect(range._maxColumnNumber).toBe(3);
            expect(range._numRows).toBe(3);
            expect(range._numColumns).toBe(2);

            range._startCell = endCell;
            range._endCell = startCell;
            expect(range._minRowNumber).toBe(3);
            expect(range._maxRowNumber).toBe(5);
            expect(range._minColumnNumber).toBe(2);
            expect(range._maxColumnNumber).toBe(3);
            expect(range._numRows).toBe(3);
            expect(range._numColumns).toBe(2);
        });
    });
});
