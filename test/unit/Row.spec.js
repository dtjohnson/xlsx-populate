"use strict";

const _ = require('lodash');
const proxyquire = require("proxyquire");

describe("Row", () => {
    let Row, Cell, row, rowNode, sheet, style, styleSheet, workbook, horizontalPageBreaks;

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
        Cell.prototype.style = jasmine.createSpy('style');

        Row = proxyquire("../../lib/Row", {
            './Cell': Cell,
            '@noCallThru': true
        });

        const Style = class {};
        if (!Style.name) Style.name = "Style";
        Style.prototype.id = jasmine.createSpy("Style.id").and.returnValue("STYLE_ID");
        Style.prototype.style = jasmine.createSpy("Style.style").and.callFake(name => `STYLE:${name}`);
        style = new Style();

        styleSheet = jasmine.createSpyObj("styleSheet", ["createStyle"]);
        styleSheet.createStyle.and.returnValue(style);

        workbook = jasmine.createSpyObj("workbook", ["sharedStrings", "styleSheet"]);
        workbook.styleSheet.and.returnValue(styleSheet);

        horizontalPageBreaks = jasmine.createSpyObj("horizontalPageBreaks", ["add"]);

        sheet = jasmine.createSpyObj('sheet', ['name', 'workbook', 'existingColumnStyleId', 'forEachExistingColumnNumber', 'horizontalPageBreaks']);
        sheet.name.and.returnValue('NAME');
        sheet.workbook.and.returnValue(workbook);
        sheet.horizontalPageBreaks.and.returnValue(horizontalPageBreaks);
        sheet.existingColumnStyleId.and.callFake(columnNumber => columnNumber === 4 ? "STYLE_ID" : undefined);
        sheet.forEachExistingColumnNumber.and.callFake(callback => _.forEach([1, 2, 4], callback));

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

    /* PUBLIC */

    describe("address", () => {
        it("should return the address", () => {
            expect(row.address()).toBe('7:7');
            expect(row.address({ anchored: true })).toBe('$7:$7');
            expect(row.address({ includeSheetName: true })).toBe("'NAME'!7:7");
            expect(row.address({ includeSheetName: true, anchored: true })).toBe("'NAME'!$7:$7");
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
            const cell = row.cell(5);
            expect(cell).toEqual(jasmine.any(Cell));
            expect(Cell).toHaveBeenCalledWith(row, 5, undefined);
            expect(row._cells[5]).toBe(cell);
        });

        it("should create a new cell as needed", () => {
            const cell = row.cell('C');
            expect(cell).toEqual(jasmine.any(Cell));
            expect(Cell).toHaveBeenCalledWith(row, 3, undefined);
            expect(row._cells[3]).toBe(cell);
        });

        it("should create a new cell with an existing column style id", () => {
            sheet.existingColumnStyleId.and.returnValue(5);
            expect(row.cell('C')).toEqual(jasmine.any(Cell));
            expect(Cell).toHaveBeenCalledWith(row, 3, 5);
        });

        it("should create a new cell with an existing row style id", () => {
            rowNode.attributes.s = 3;
            expect(row.cell('C')).toEqual(jasmine.any(Cell));
            expect(Cell).toHaveBeenCalledWith(row, 3, 3);
        });

        it("should create a new cell with an existing row and column style id", () => {
            sheet.existingColumnStyleId.and.returnValue(5);
            rowNode.attributes.s = 3;
            expect(row.cell('C')).toEqual(jasmine.any(Cell));
            expect(Cell).toHaveBeenCalledWith(row, 3, 3);
        });

        it("should throw an exception on an index of 0", () => {
            expect(() => row.cell(0)).toThrowError(RangeError);
        });

        it("should throw an exception on an index of -1", () => {
            expect(() => row.cell(-1)).toThrowError(RangeError);
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

    describe("style", () => {
        beforeEach(() => {
            spyOn(row, "_createStyleIfNeeded");
            row._style = style;
        });

        it("should get a single style", () => {
            expect(row.style("foo")).toBe("STYLE:foo");
            expect(style.style).toHaveBeenCalledWith("foo");
            expect(row._createStyleIfNeeded).toHaveBeenCalledWith();
        });

        it("should get multiple styles", () => {
            expect(row.style(["foo", "bar", "baz"])).toEqualJson({
                foo: "STYLE:foo", bar: "STYLE:bar", baz: "STYLE:baz"
            });
            expect(style.style).toHaveBeenCalledWith("foo");
            expect(style.style).toHaveBeenCalledWith("bar");
            expect(style.style).toHaveBeenCalledWith("baz");
            expect(row._createStyleIfNeeded).toHaveBeenCalledWith();
        });

        it("should set a single style", () => {
            expect(row._cells[2]).toBeDefined();
            expect(row._cells[4]).toBeUndefined();

            expect(row.style("foo", "value")).toBe(row);
            expect(style.style).toHaveBeenCalledWith("foo", "value");

            expect(row._cells[1]).toBeUndefined();
            expect(row._cells[2].style).toHaveBeenCalledWith("foo", "value");
            expect(row._cells[3]).toBeUndefined();
            expect(row._cells[4].style).toHaveBeenCalledWith("foo", "value");
        });

        it("should assign a style when asked", () => {
            row._style = undefined;
            expect(row._cells[2]).toBeDefined();
            expect(row._cells[4]).toBeUndefined();

            expect(row.style(style)).toBe(row);

            expect(row._cells[1]).toBeUndefined();
            expect(row._cells[2].style).toHaveBeenCalledWith(style);
            expect(row._cells[3]).toBeUndefined();
            expect(row._cells[4].style).toHaveBeenCalledWith(style);
        });

        it("should set multiple styles", () => {
            expect(row.style({
                foo: "FOO", bar: "BAR", baz: "BAZ"
            })).toBe(row);
            expect(style.style).toHaveBeenCalledWith("foo", "FOO");
            expect(style.style).toHaveBeenCalledWith("bar", "BAR");
            expect(style.style).toHaveBeenCalledWith("baz", "BAZ");

            expect(row._cells[2].style).toHaveBeenCalledWith("foo", "FOO");
            expect(row._cells[2].style).toHaveBeenCalledWith("bar", "BAR");
            expect(row._cells[2].style).toHaveBeenCalledWith("baz", "BAZ");
            expect(row._cells[4].style).toHaveBeenCalledWith("foo", "FOO");
            expect(row._cells[4].style).toHaveBeenCalledWith("bar", "BAR");
            expect(row._cells[4].style).toHaveBeenCalledWith("baz", "BAZ");
        });
    });

    describe("workbook", () => {
        it("should return the workbook", () => {
            expect(row.workbook()).toBe(workbook);
        });
    });

    describe('addPageBreak', () => {
        it("should add a rowBreak and return the row", () => {
            expect(row.addPageBreak()).toBe(row);
        });
    });

    /* INTERNAL */

    describe("clearCellsUsingSharedFormula", () => {
        it("should clear cells with matching shared formula", () => {
            row._cells = [
                undefined,
                {
                    sharesFormula: jasmine.createSpy("sharesFormula").and.returnValue(true),
                    clear: jasmine.createSpy("clear")
                },
                undefined,
                {
                    sharesFormula: jasmine.createSpy("sharesFormula").and.returnValue(false),
                    clear: jasmine.createSpy("clear")
                }
            ];

            row.clearCellsUsingSharedFormula(7);
            expect(row._cells[1].sharesFormula).toHaveBeenCalledWith(7);
            expect(row._cells[1].clear).toHaveBeenCalledWith();
            expect(row._cells[3].sharesFormula).toHaveBeenCalledWith(7);
            expect(row._cells[3].clear).not.toHaveBeenCalled();
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

    describe("hasCell", () => {
        it("should return true/false if the cell exists or not", () => {
            expect(row.hasCell(1)).toBe(false);
            expect(row.hasCell(2)).toBe(true);
            expect(row.hasCell(3)).toBe(false);
        });

        it("should throw an exception on an index of 0", () => {
            expect(() => row.hasCell(0)).toThrowError(RangeError);
        });

        it("should throw an exception on an index of -1", () => {
            expect(() => row.hasCell(-1)).toThrowError(RangeError);
        });
    });

    describe("hasStyle", () => {
        it("should return true/false if the row has a style set", () => {
            expect(row.hasStyle()).toBe(false);
            rowNode.attributes.s = 0;
            expect(row.hasStyle()).toBe(true);
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

    describe("toXml", () => {
        it("should return the node", () => {
            expect(row.toXml()).toBe(rowNode);
        });
    });

    /* PRIVATE */

    describe("_createStyleIfNeeded", () => {
        it("should create a style", () => {
            rowNode.attributes.s = 3;
            row._createStyleIfNeeded();
            expect(row._style).toBe(style);
            expect(rowNode.attributes.s).toBe("STYLE_ID");
            expect(styleSheet.createStyle).toHaveBeenCalledWith(3);
        });

        it("should NOT create a style", () => {
            const existingStyle = {};
            row._style = existingStyle;
            row._createStyleIfNeeded();
            expect(row._style).toBe(existingStyle);
            expect(styleSheet.createStyle).not.toHaveBeenCalled();
        });
    });

    describe("_init", () => {
        it("should store existing rows", () => {
            expect(row._cells).toEqual([undefined, undefined, jasmine.any(Cell)]);
            expect(rowNode.children).toBe(row._cells);
            expect(Cell).toHaveBeenCalledWith(row, {
                name: 'c',
                attributes: {
                    r: "B7"
                }
            });
        });
    });
});
