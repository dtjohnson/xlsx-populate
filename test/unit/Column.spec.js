"use strict";

const _ = require("lodash");
const proxyquire = require("proxyquire");

describe("Column", () => {
    let Column, column, columnNode, sheet, style, styleSheet, workbook, existingRows, verticalPageBreaks;

    beforeEach(() => {
        Column = proxyquire("../../lib/Column", {
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

        verticalPageBreaks = jasmine.createSpyObj("verticalPageBreaks", ["add"]);

        existingRows = [
            {
                hasStyle: () => false,
                hasCell: () => false,
                cell: jasmine.createSpy("cell[0]")
            },
            {
                hasStyle: () => true,
                hasCell: () => false,
                _cell: jasmine.createSpyObj("cell[1]", ["style"]),
                cell: jasmine.createSpy("cell[1]").and.callFake(function () {
                    return this._cell;
                })
            },
            {
                hasStyle: () => false,
                hasCell: () => true,
                _cell: jasmine.createSpyObj("cell[2]", ["style"]),
                cell: jasmine.createSpy("cell[2]").and.callFake(function () {
                    return this._cell;
                })
            }
        ];

        sheet = jasmine.createSpyObj('sheet', ['cell', 'name', 'workbook', 'forEachExistingRow', 'verticalPageBreaks']);
        sheet.cell.and.returnValue('CELL');
        sheet.name.and.returnValue('NAME');
        sheet.workbook.and.returnValue(workbook);
        sheet.verticalPageBreaks.and.returnValue(verticalPageBreaks);
        sheet.forEachExistingRow.and.callFake(callback => _.forEach(existingRows, callback));

        columnNode = {
            name: 'col',
            attributes: {
                min: 5,
                max: 5
            }
        };

        column = new Column(sheet, columnNode);
    });

    /* PUBLIC */

    describe("address", () => {
        it("should return the address", () => {
            expect(column.address()).toBe('E:E');
            expect(column.address({ anchored: true })).toBe('$E:$E');
            expect(column.address({ includeSheetName: true })).toBe("'NAME'!E:E");
            expect(column.address({ includeSheetName: true, anchored: true })).toBe("'NAME'!$E:$E");
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

    describe("style", () => {
        beforeEach(() => {
            spyOn(column, "_createStyleIfNeeded");
            column._style = style;
        });

        it("should get a single style", () => {
            expect(column.style("foo")).toBe("STYLE:foo");
            expect(style.style).toHaveBeenCalledWith("foo");
            expect(column._createStyleIfNeeded).toHaveBeenCalledWith();
        });

        it("should get multiple styles", () => {
            expect(column.style(["foo", "bar", "baz"])).toEqualJson({
                foo: "STYLE:foo", bar: "STYLE:bar", baz: "STYLE:baz"
            });
            expect(style.style).toHaveBeenCalledWith("foo");
            expect(style.style).toHaveBeenCalledWith("bar");
            expect(style.style).toHaveBeenCalledWith("baz");
            expect(column._createStyleIfNeeded).toHaveBeenCalledWith();
        });

        it("should set a single style", () => {
            expect(column.style("foo", "value")).toBe(column);
            expect(style.style).toHaveBeenCalledWith("foo", "value");

            expect(existingRows[0].cell).not.toHaveBeenCalled();
            expect(existingRows[1].cell).toHaveBeenCalledWith(5);
            expect(existingRows[1]._cell.style).toHaveBeenCalledWith("foo", "value");
            expect(existingRows[2].cell).toHaveBeenCalledWith(5);
            expect(existingRows[2]._cell.style).toHaveBeenCalledWith("foo", "value");
        });

        it("should assign a style when asked", () => {
            column._style = undefined;
            column.style(style);
            expect(column._style).toBe(style);
            expect(column._node.attributes.style).toBe(style.id());

            expect(existingRows[0].cell).not.toHaveBeenCalled();
            expect(existingRows[1].cell).toHaveBeenCalledWith(5);
            expect(existingRows[1]._cell.style).toHaveBeenCalledWith(style);
            expect(existingRows[2].cell).toHaveBeenCalledWith(5);
            expect(existingRows[2]._cell.style).toHaveBeenCalledWith(style);
        });

        it("should set multiple styles", () => {
            expect(column.style({
                foo: "FOO", bar: "BAR", baz: "BAZ"
            })).toBe(column);
            expect(style.style).toHaveBeenCalledWith("foo", "FOO");
            expect(style.style).toHaveBeenCalledWith("bar", "BAR");
            expect(style.style).toHaveBeenCalledWith("baz", "BAZ");

            expect(existingRows[0].cell).not.toHaveBeenCalled();
            expect(existingRows[1].cell).toHaveBeenCalledWith(5);
            expect(existingRows[1]._cell.style).toHaveBeenCalledWith("foo", "FOO");
            expect(existingRows[1]._cell.style).toHaveBeenCalledWith("bar", "BAR");
            expect(existingRows[1]._cell.style).toHaveBeenCalledWith("baz", "BAZ");
            expect(existingRows[2].cell).toHaveBeenCalledWith(5);
            expect(existingRows[2]._cell.style).toHaveBeenCalledWith("foo", "FOO");
            expect(existingRows[2]._cell.style).toHaveBeenCalledWith("bar", "BAR");
            expect(existingRows[2]._cell.style).toHaveBeenCalledWith("baz", "BAZ");
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
            expect(column.workbook()).toBe(workbook);
        });
    });

    describe('addPageBreak', () => {
        it("should add a colBreak and return the column", () => {
            expect(column.addPageBreak()).toBe(column);
        });
    });

    /* INTERNAL */

    describe("toXml", () => {
        it("should return the object representation", () => {
            expect(column.toXml()).toBe(columnNode);
        });
    });

    /* PRIVATE */

    describe("_createStyleIfNeeded", () => {
        it("should create a style", () => {
            spyOn(column, "width");
            columnNode.attributes.style = 3;
            column._createStyleIfNeeded();
            expect(column._style).toBe(style);
            expect(columnNode.attributes.style).toBe("STYLE_ID");
            expect(styleSheet.createStyle).toHaveBeenCalledWith(3);
            expect(column.width).toHaveBeenCalledWith(9.140625);
        });

        it("should NOT create a style", () => {
            const existingStyle = {};
            column._style = existingStyle;
            column._createStyleIfNeeded();
            expect(column._style).toBe(existingStyle);
            expect(styleSheet.createStyle).not.toHaveBeenCalled();
        });
    });
});
