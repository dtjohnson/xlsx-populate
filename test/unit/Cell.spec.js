"use strict";

const proxyquire = require("proxyquire");

describe("Cell", () => {
    let Cell, cell, cellNode, row, sheet, workbook, sharedStrings, styleSheet, style, FormulaError;

    beforeEach(() => {
        FormulaError = jasmine.createSpyObj("FormulaError", ["getError"]);
        FormulaError.getError.and.returnValue("ERROR");

        Cell = proxyquire("../../lib/Cell", {
            './FormulaError': FormulaError,
            '@noCallThru': true
        });

        style = jasmine.createSpyObj("style", ["style", "id"]);
        style.id.and.returnValue(4);
        style.style.and.callFake(name => `STYLE:${name}`);

        styleSheet = jasmine.createSpyObj("styleSheet", ["createStyle"]);
        styleSheet.createStyle.and.returnValue(style);

        sharedStrings = jasmine.createSpyObj("sharedStrings", ['getIndexForString', 'getStringByIndex']);
        sharedStrings.getIndexForString.and.returnValue(7);
        sharedStrings.getStringByIndex.and.returnValue("STRING");

        workbook = jasmine.createSpyObj("workbook", ["sharedStrings", "styleSheet"]);
        workbook.sharedStrings.and.returnValue(sharedStrings);
        workbook.styleSheet.and.returnValue(styleSheet);

        sheet = jasmine.createSpyObj('sheet', ['createStyle', 'activeCell', 'updateMaxSharedFormulaId', 'name', 'column', 'clearCellsUsingSharedFormula', 'cell', 'range', 'hyperlink']);
        sheet.activeCell.and.returnValue("ACTIVE CELL");
        sheet.name.and.returnValue("NAME");
        sheet.column.and.returnValue("COLUMN");
        sheet.hyperlink.and.returnValue("HYPERLINK");
        sheet.range.and.returnValue("RANGE");

        row = jasmine.createSpyObj('row', ['sheet', 'workbook']);
        row.sheet.and.returnValue(sheet);
        row.workbook.and.returnValue(workbook);

        cellNode = {
            name: 'c',
            attributes: {
                r: "C7"
            },
            children: []
        };

        cell = new Cell(row, cellNode);
    });

    describe("active", () => {
        it("should return true/false", () => {
            expect(cell.active()).toBe(false);
            sheet.activeCell.and.returnValue(cell);
            expect(cell.active()).toBe(true);
        });

        it("should set the sheet active cell", () => {
            expect(cell.active(true)).toBe(cell);
            expect(sheet.activeCell).toHaveBeenCalledWith(cell);
        });

        it("should throw an error if attempting to deactivate", () => {
            expect(() => cell.active(false)).toThrow();
        });
    });

    describe("address", () => {
        it("should return the address", () => {
            expect(cell.address()).toBe('C7');
            expect(cell.address({ rowAnchored: true })).toBe('C$7');
            expect(cell.address({ columnAnchored: true })).toBe('$C7');
            expect(cell.address({ includeSheetName: true })).toBe("'NAME'!C7");
            expect(cell.address({ includeSheetName: true, rowAnchored: true, columnAnchored: true })).toBe("'NAME'!$C$7");
            expect(cell.address({ anchored: true })).toBe("$C$7");
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

    describe("hyperlink", () => {
        it("should get the hyperlink from the sheet", () => {
            expect(cell.hyperlink()).toBe("HYPERLINK");
            expect(sheet.hyperlink).toHaveBeenCalledWith("C7");
        });

        it("should set the hyperlink on the sheet", () => {
            expect(cell.hyperlink("HYPERLINK")).toBe(cell);
            expect(sheet.hyperlink).toHaveBeenCalledWith("C7", "HYPERLINK");
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

    describe("rangeTo", () => {
        it("should create a range", () => {
            expect(cell.rangeTo("OTHER")).toBe("RANGE");
            expect(sheet.range).toHaveBeenCalledWith(cell, "OTHER");
        });
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
        it("should create a new style without a style ID", () => {
            expect(cell._style).toBeUndefined();
            cell.style("foo");
            expect(styleSheet.createStyle).toHaveBeenCalledWith(undefined);
            expect(cell._style).toBe(style);
            expect(cellNode.attributes.s).toBe(4);
        });

        it("should create a new style with a style ID", () => {
            cellNode.attributes.s = 2;
            expect(cell._style).toBeUndefined();
            cell.style("foo");
            expect(styleSheet.createStyle).toHaveBeenCalledWith(2);
            expect(cell._style).toBe(style);
            expect(cellNode.attributes.s).toBe(4);
        });

        it("should not create a style if one already exists", () => {
            cell._style = style;
            cell.style("foo");
            expect(styleSheet.createStyle).not.toHaveBeenCalled();
        });

        it("should get a single style", () => {
            expect(cell.style("foo")).toBe("STYLE:foo");
            expect(style.style).toHaveBeenCalledWith("foo");
        });

        it("should get multiple styles", () => {
            expect(cell.style(["foo", "bar", "baz"])).toEqualJson({
                foo: "STYLE:foo", bar: "STYLE:bar", baz: "STYLE:baz"
            });
            expect(style.style).toHaveBeenCalledWith("foo");
            expect(style.style).toHaveBeenCalledWith("bar");
            expect(style.style).toHaveBeenCalledWith("baz");
        });

        it("should set a single style", () => {
            expect(cell.style("foo", "value")).toBe(cell);
            expect(style.style).toHaveBeenCalledWith("foo", "value");
        });

        it("should set multiple styles", () => {
            expect(cell.style({
                foo: "FOO", bar: "BAR", baz: "BAZ"
            })).toBe(cell);
            expect(style.style).toHaveBeenCalledWith("foo", "FOO");
            expect(style.style).toHaveBeenCalledWith("bar", "BAR");
            expect(style.style).toHaveBeenCalledWith("baz", "BAZ");
        });
    });

    describe("value", () => {
        it("should get/set the value", () => {
            expect(cell.value()).toBeUndefined();

            cell.value(5);
            expect(cell.value()).toBe(5);
            expect(cellNode.attributes.t).toBeUndefined();
            expect(cellNode.children).toEqualJson([{
                name: 'v',
                children: [5]
            }]);

            cell.value(-3.7);
            expect(cell.value()).toBe(-3.7);
            expect(cellNode.attributes.t).toBeUndefined();
            expect(cellNode.children).toEqualJson([{
                name: 'v',
                children: [-3.7]
            }]);

            cell.value(true);
            expect(cell.value()).toBe(true);
            expect(cellNode.attributes.t).toBe('b');
            expect(cellNode.children).toEqualJson([{
                name: 'v',
                children: [1]
            }]);

            cell.value(false);
            expect(cell.value()).toBe(false);
            expect(cellNode.attributes.t).toBe('b');
            expect(cellNode.children).toEqualJson([{
                name: 'v',
                children: [0]
            }]);

            cell.value(new Date(2017, 0, 1));
            expect(cell.value()).toBe(42736);
            expect(cellNode.attributes.t).toBeUndefined();
            expect(cellNode.children).toEqualJson([{
                name: 'v',
                children: [42736]
            }]);

            cell.value("some string");
            expect(cell.value()).toBe("STRING");
            expect(cellNode.attributes.t).toBe('s');
            expect(cellNode.children).toEqualJson([{
                name: 'v',
                children: [7]
            }]);
            expect(sharedStrings.getStringByIndex).toHaveBeenCalledWith(7);
            expect(sharedStrings.getIndexForString).toHaveBeenCalledWith("some string");

            cellNode.attributes.t = "inlineStr";
            cellNode.children = [{
                name: 'is',
                attributes: {},
                children: [{
                    name: 't',
                    attributes: {},
                    children: ["inline string"]
                }]
            }];
            expect(cell.value()).toBe("inline string");

            cellNode.attributes.t = "e";
            cellNode.children = [{
                name: 'v',
                children: ["#ERR!"]
            }];
            expect(cell.value()).toEqual("ERROR");
            expect(FormulaError.getError).toHaveBeenCalledWith("#ERR!");

            cell.value(undefined);
            expect(cell.value()).toBeUndefined();
            expect(cellNode.attributes.t).toBeUndefined();
            expect(cellNode.children).toEqualJson([]);
        });
    });

    describe("workbook", () => {
        it("should return the workbook from the row", () => {
            expect(cell.workbook()).toBe(workbook);
        });
    });

    describe("getSharedRefFormula", () => {
        it("should return undefined if not a formula", () => {
            expect(cell.getSharedRefFormula()).toBeUndefined();
        });

        it("should return undefined if not a shared formula ref", () => {
            cellNode.children = [{
                name: 'f',
                attributes: {}
            }];
            expect(cell.getSharedRefFormula()).toBeUndefined();
        });

        it("should return the shared formula", () => {
            cellNode.children = [{
                name: 'f',
                attributes: { ref: "A1:C3", si: 5 },
                children: ["FORMULA"]
            }];
            expect(cell.getSharedRefFormula()).toBe("FORMULA");
        });
    });

    describe("sharesFormula", () => {
        it("should return true/false if shares a given formula or not", () => {
            cellNode.children = [{
                name: 'f',
                attributes: {
                    si: 6
                }
            }];

            expect(cell.sharesFormula(6)).toBe(true);
            expect(cell.sharesFormula(3)).toBe(false);
        });

        it("should return undefined if doesn't share any formula", () => {
            expect(cell.sharesFormula(6)).toBeUndefined();
        });
    });

    describe("setSharedFormula", () => {
        it("should set a ref shared formula", () => {
            cell.setSharedFormula(3, "A1*A2", "A1:C1");
            expect(cellNode.children).toEqualJson([{
                name: 'f',
                attributes: {
                    t: "shared",
                    si: 3,
                    ref: "A1:C1"
                },
                children: ["A1*A2"]
            }]);
        });

        it("should set a dependent shared formula", () => {
            cell.setSharedFormula(3);
            expect(cellNode.children).toEqualJson([{
                name: 'f',
                attributes: {
                    t: "shared",
                    si: 3
                },
                children: []
            }]);
        });
    });

    describe("toObject", () => {
        it("should return the node with stored formula values cleared", () => {
            cell._node.children = [{
                name: 'f',
                attributes: {},
                children: []
            }, {
                name: 'v',
                attributes: {},
                children: ["VALUE"]
            }];

            expect(cell.toXmls()).toEqualJson({
                name: "c",
                attributes: {
                    r: "C7"
                },
                children: [{
                    name: 'f',
                    attributes: {},
                    children: []
                }]
            });
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

    describe("_init", () => {
        it("should parse the address", () => {
            cell._init(cellNode);
            expect(cell._ref).toEqualJson({
                type: 'cell',
                columnName: 'C',
                columnNumber: 3,
                columnAnchored: false,
                rowNumber: 7,
                rowAnchored: false
            });
        });

        it("should update the sheet max shared formula ID", () => {
            spyOn(cell, '_getSharedFormulaRefId').and.returnValue("REF_ID");
            cell._init(cellNode);
            expect(sheet.updateMaxSharedFormulaId).toHaveBeenCalledWith("REF_ID");
        });
    });
});
