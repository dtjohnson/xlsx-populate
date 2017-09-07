"use strict";

const proxyquire = require("proxyquire");

describe("Sheet", () => {
    let Sheet, Range, Row, Cell, Column, Relationships, sheet, idNode, sheetNode, workbook;

    beforeEach(() => {
        let i = 0;
        workbook = jasmine.createSpyObj("workbook", ["scopedDefinedName", "activeSheet", "sheets", "deleteSheet", "moveSheet"]);
        workbook.scopedDefinedName.and.returnValue("DEFINED NAME");
        workbook.activeSheet.and.returnValue("ACTIVE SHEET");

        Range = jasmine.createSpy("Range");
        Row = jasmine.createSpy("Row");
        Row.prototype.rowNumber = jasmine.createSpy().and.callFake(() => ++i);
        Row.prototype.find = jasmine.createSpy('find');
        Column = jasmine.createSpy("Column");
        Cell = jasmine.createSpy("Cell");
        Cell.prototype.address = jasmine.createSpy("Cell.address").and.returnValue("ADDRESS");

        Relationships = jasmine.createSpy("Relationships");
        Relationships.prototype.findById = jasmine.createSpy("Relationships.findById").and.callFake(id => ({ attributes: { Target: `TARGET:${id}` } }));
        Relationships.prototype.add = jasmine.createSpy("Relationships.add").and.returnValue({ attributes: { Id: "ID" } });

        Sheet = proxyquire("../../lib/Sheet", {
            './Range': Range,
            './Row': Row,
            './Column': Column,
            './Cell': Cell,
            './Relationships': Relationships,
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

    describe("active", () => {
        it("should return true/false", () => {
            expect(sheet.active()).toBe(false);
            workbook.activeSheet.and.returnValue(sheet);
            expect(sheet.active()).toBe(true);
        });

        it("should set the workbook active sheet", () => {
            expect(sheet.active(true)).toBe(sheet);
            expect(workbook.activeSheet).toHaveBeenCalledWith(sheet);
        });

        it("should throw an error if attempting to deactivate", () => {
            expect(() => sheet.active(false)).toThrow();
        });
    });

    describe("activeCell", () => {
        let cell;
        beforeEach(() => {
            cell = new Cell();
            spyOn(sheet, 'cell').and.returnValue(cell);
        });

        it("should get the default active cell", () => {
            expect(sheet.activeCell()).toBe(cell);
            expect(sheet.cell).toHaveBeenCalledWith("A1");
        });

        it("should get the active cell", () => {
            sheetNode.children.push({
                name: "sheetViews",
                attributes: {},
                children: [{
                    name: "sheetView",
                    attributes: {
                        workbookViewId: 0
                    },
                    children: [{
                        name: "selection",
                        attributes: {
                            activeCell: "B5"
                        }
                    }]
                }]
            });

            expect(sheet.activeCell()).toBe(cell);
            expect(sheet.cell).toHaveBeenCalledWith("B5");
        });

        it("should set the active cell by cell", () => {
            expect(sheet.activeCell(cell)).toBe(sheet);
            expect(sheetNode.children[1]).toEqualJson({
                name: "sheetViews",
                attributes: {},
                children: [{
                    name: "sheetView",
                    attributes: {
                        workbookViewId: 0
                    },
                    children: [{
                        name: "selection",
                        attributes: {
                            activeCell: "ADDRESS",
                            sqref: "ADDRESS"
                        },
                        children: []
                    }]
                }]
            });

            expect(sheet.cell).not.toHaveBeenCalled();
        });

        it("should set the active cell by address", () => {
            expect(sheet.activeCell("C6")).toBe(sheet);
            expect(sheetNode.children[1].children[0].children[0].attributes).toEqualJson({
                activeCell: "ADDRESS",
                sqref: "ADDRESS"
            });

            expect(sheet.cell).toHaveBeenCalledWith("C6");
        });

        it("should set the active cell by row and column", () => {
            expect(sheet.activeCell(5, 4)).toBe(sheet);
            expect(sheetNode.children[1].children[0].children[0].attributes).toEqualJson({
                activeCell: "ADDRESS",
                sqref: "ADDRESS"
            });

            expect(sheet.cell).toHaveBeenCalledWith(5, 4);
        });
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
            const colNode = Column.calls.argsFor(0)[1];
            expect(colNode).toEqualJson({
                name: 'col',
                attributes: {
                    min: 5,
                    max: 5
                },
                children: []
            });
            expect(Column).toHaveBeenCalledWith(sheet, colNode);
            expect(sheet._colNodes[5]).toEqualJson(colNode);
        });

        it("should break an existing column", () => {
            const existingColNode = {
                name: "col",
                attributes: {
                    min: 4,
                    max: 7
                },
                children: []
            };

            sheet._colNodes[4] = sheet._colNodes[5] = sheet._colNodes[6] = sheet._colNodes[7] = existingColNode;

            const column = sheet.column('F');
            expect(column).toEqual(jasmine.any(Column));
            expect(sheet._columns[6]).toBe(column);
            const colNode = Column.calls.argsFor(0)[1];
            expect(colNode).toEqualJson({
                name: 'col',
                attributes: {
                    min: 6,
                    max: 6
                },
                children: []
            });
            expect(Column).toHaveBeenCalledWith(sheet, colNode);
            expect(sheet._colNodes).toEqualJson([
                null,
                null,
                null,
                null,
                {
                    name: "col",
                    attributes: {
                        min: 4,
                        max: 5
                    },
                    children: []
                },
                {
                    name: "col",
                    attributes: {
                        min: 4,
                        max: 5
                    },
                    children: []
                },
                colNode,
                {
                    name: "col",
                    attributes: {
                        min: 7,
                        max: 7
                    },
                    children: []
                }
            ]);
        });
    });

    describe("definedName", () => {
        it("should return the defined name", () => {
            expect(sheet.definedName("FOO")).toBe("DEFINED NAME");
            expect(workbook.scopedDefinedName).toHaveBeenCalledWith(sheet, "FOO");
        });

        it("should set the defined name", () => {
            expect(sheet.definedName("NAME", "REF")).toBe(sheet);
            expect(workbook.scopedDefinedName).toHaveBeenCalledWith(sheet, "NAME", "REF");
        });
    });

    describe("delete", () => {
        it("should call the workbook delete method", () => {
            expect(sheet.delete()).toBe(workbook);
            expect(workbook.deleteSheet).toHaveBeenCalledWith(sheet);
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

    describe("hidden", () => {
        it("should return the hidden state", () => {
            expect(sheet.hidden()).toBe(false);

            idNode.attributes.state = "hidden";
            expect(sheet.hidden()).toBe(true);

            idNode.attributes.state = "veryHidden";
            expect(sheet.hidden()).toBe('very');
        });

        it("should hide/unhide the sheet", () => {
            workbook.sheets.and.returnValue([
                sheet,
                {
                    hidden: jasmine.createSpy("hidden").and.returnValue(false)
                }
            ]);

            expect(sheet.hidden(true)).toBe(sheet);
            expect(idNode.attributes.state).toBe("hidden");

            expect(sheet.hidden('very')).toBe(sheet);
            expect(idNode.attributes.state).toBe("veryHidden");

            expect(sheet.hidden(false)).toBe(sheet);
            expect(idNode.attributes.state).toBeUndefined();
        });

        it("should hide the sheet and activate a different one", () => {
            const otherSheet = {
                active: jasmine.createSpy("active").and.returnValue(false),
                hidden: jasmine.createSpy("hidden").and.returnValue(false)
            };
            workbook.sheets.and.returnValue([sheet, otherSheet]);

            spyOn(sheet, "active").and.returnValue(true);
            sheet.hidden(true);
            expect(otherSheet.active).toHaveBeenCalledWith(true);
        });

        it("should throw an error if trying to hide the only visible sheet", () => {
            workbook.sheets.and.returnValue([
                sheet,
                {
                    hidden: jasmine.createSpy("hidden").and.returnValue(true)
                }
            ]);

            expect(() => sheet.hidden(true)).toThrow();
        });
    });

    describe("move", () => {
        it("should call the workbook move method", () => {
            expect(sheet.move("BEFORE")).toBe(sheet);
            expect(workbook.moveSheet).toHaveBeenCalledWith(sheet, "BEFORE");
        });
    });

    describe("name", () => {
        it("should return the sheet name", () => {
            expect(sheet.name()).toBe("SHEET NAME");
        });

        it("should set the sheet name", () => {
            expect(sheet.name("a new name")).toBe(sheet);
            expect(sheet.name()).toBe("a new name");
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

    describe("tabColor", () => {
        it("should get the tab color", () => {
            expect(sheet.tabColor()).toBeUndefined();

            sheet._sheetPrNode.children = [{
                name: "tabColor",
                attributes: {
                    rgb: "RGB"
                }
            }];
            expect(sheet.tabColor()).toEqualJson({
                rgb: "RGB"
            });

            sheet._sheetPrNode.children = [{
                name: "tabColor",
                attributes: {
                    theme: 0
                }
            }];
            expect(sheet.tabColor()).toEqualJson({
                theme: 0
            });

            sheet._sheetPrNode.children = [{
                name: "tabColor",
                attributes: {
                    rgb: "RGB",
                    tint: "TINT"
                }
            }];
            expect(sheet.tabColor()).toEqualJson({
                rgb: "RGB",
                tint: "TINT"
            });

            sheet._sheetPrNode.children = [{
                name: "tabColor",
                attributes: {
                    indexed: 5
                }
            }];
            expect(sheet.tabColor()).toEqualJson({
                rgb: "FFFF00"
            });
        });

        it("should set the tab color", () => {
            expect(sheet.tabColor("ff0000")).toBe(sheet);
            expect(sheet._sheetPrNode.children).toEqualJson([{
                name: "tabColor",
                attributes: {
                    rgb: "FF0000"
                },
                children: []
            }]);

            expect(sheet.tabColor(5)).toBe(sheet);
            expect(sheet._sheetPrNode.children).toEqualJson([{
                name: "tabColor",
                attributes: {
                    theme: 5
                },
                children: []
            }]);

            expect(sheet.tabColor({ rgb: "ff0000", tint: -0.5 })).toBe(sheet);
            expect(sheet._sheetPrNode.children).toEqualJson([{
                name: "tabColor",
                attributes: {
                    rgb: "FF0000",
                    tint: -0.5
                },
                children: []
            }]);

            expect(sheet.tabColor(null)).toBe(sheet);
            expect(sheet._sheetPrNode.children).toEqualJson([]);
        });
    });

    describe("tabSelected", () => {
        let sheetViewNode;

        beforeEach(() => {
            sheetViewNode = { attributes: {} };
            spyOn(sheet, "_getOrCreateSheetViewNode").and.returnValue(sheetViewNode);
        });

        it("should return the tab selected state", () => {
            expect(sheet.tabSelected()).toBe(false);

            sheetViewNode.attributes.tabSelected = 1;
            expect(sheet.tabSelected()).toBe(true);
        });

        it("should select/deselect the sheet tab", () => {
            expect(sheet.tabSelected(true)).toBe(sheet);
            expect(sheetViewNode.attributes.tabSelected).toBe(1);

            expect(sheet.tabSelected(false)).toBe(sheet);
            expect(sheetViewNode.attributes.tabSelected).toBeUndefined();
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
            sheet._colNodes[5] = {
                attributes: {
                    style: "STYLE ID"
                }
            };
            expect(sheet.existingColumnStyleId(5)).toBe("STYLE ID");
        });
    });

    describe("forEachExistingColumnNumber", () => {
        it("should call the callback for each existing column number", () => {
            sheet._colNodes = [null, "NODE1", "NODE2"];
            const callback = jasmine.createSpy("callback");
            sheet.forEachExistingColumnNumber(callback);
            expect(callback.calls.count()).toBe(2);
            expect(callback).toHaveBeenCalledWith(1);
            expect(callback).toHaveBeenCalledWith(2);
        });
    });

    describe("forEachExistingRow", () => {
        it("should call the callback for each existing row", () => {
            sheet._rows = [null, "ROW1", "ROW2"];
            const callback = jasmine.createSpy("callback");
            sheet.forEachExistingRow(callback);
            expect(callback.calls.count()).toBe(2);
            expect(callback).toHaveBeenCalledWith("ROW1", 1);
            expect(callback).toHaveBeenCalledWith("ROW2", 2);
        });
    });

    describe("dataValidation", () => {
        it("should return the dataValidation Object", () => {
            sheet._dataValidations = {
                A1: {
                    name: 'dataValidation',
                    attributes: {
                        type: 'list',
                        sqref: 'A1'
                    },
                    children: [
                        {
                            name: 'formula1',
                            atrributes: {},
                            children: ['test1, test2, test3']
                        }
                    ]
                },
                A2: {
                    name: 'dataValidation',
                    attributes: {
                        type: 'list',
                        sqref: 'A2'
                    },
                    children: [
                        {
                            name: 'formula1',
                            atrributes: {},
                            children: ['test1, test2, test3']
                        }
                    ]
                }
            };

            expect(sheet.dataValidation("A1")).toEqualJson({
                type: 'list',
                formula1: 'test1, test2, test3'
            });
            
            expect(sheet.dataValidation("A2")).toEqualJson({
                type: 'list',
                formula1: 'test1, test2, test3'
            });

            expect(sheet.dataValidation("A3")).toBe(false);
        });

        it("should add a dataValidations entry", () => {
            expect(sheet._dataValidations).toEqualJson({});
            expect(sheet.dataValidation("A1", "TEST")).toBe(sheet);
            expect(sheet._dataValidations).toEqualJson({
                A1: {
                    name: 'dataValidation',
                    attributes: {
                        type: 'list',
                        allowBlank: false,
                        showInputMessage: false,
                        prompt: '',
                        promptTitle: '',
                        showErrorMessage: false,
                        error: '',
                        errorTitle: '',
                        operator: '',
                        sqref: 'A1'
                    },
                    children: [
                        {
                            name: 'formula1',
                            atrributes: {},
                            children: ['TEST']
                        },
                        {
                            name: 'formula2',
                            atrributes: {},
                            children: ['']
                        }
                    ]
                }
            });
        });

        it("should remove a dataValidation entry", () => {
            sheet._dataValidations = {
                ADDRESS1: {},
                ADDRESS2: {}
            };

            expect(sheet.dataValidation("ADDRESS3", false)).toBe(false);
            expect(sheet._dataValidations).toEqualJson({
                ADDRESS1: {},
                ADDRESS2: {}
            });

            expect(sheet.dataValidation("ADDRESS1", false)).toBe(true);
            expect(sheet._dataValidations).toEqualJson({
                ADDRESS2: {}
            });

            expect(sheet.dataValidation("ADDRESS2", false)).toBe(true);
            expect(sheet._dataValidations).toEqualJson({});
        });
    });

    describe("hyperlink", () => {
        it("should return the hyperlink", () => {
            sheet._hyperlinks = {
                ADDRESS1: { attributes: { 'r:id': "ID1" } },
                ADDRESS2: { attributes: { 'r:id': "ID2" } }
            };

            expect(sheet.hyperlink("ADDRESS1")).toBe("TARGET:ID1");
            expect(sheet.hyperlink("ADDRESS2")).toBe("TARGET:ID2");
            expect(sheet.hyperlink("ADDRESS3")).toBeUndefined();
        });

        it("should add a hyperlink entry", () => {
            expect(sheet._hyperlinks).toEqualJson({});
            expect(sheet.hyperlink("ADDRESS", "HYPERLINK")).toBe(sheet);
            expect(sheet._hyperlinks).toEqualJson({
                ADDRESS: {
                    name: 'hyperlink',
                    attributes: {
                        'r:id': "ID",
                        ref: "ADDRESS"
                    },
                    children: []
                }
            });
        });

        it("should remove a hyperlink entry", () => {
            sheet._hyperlinks = {
                ADDRESS1: {},
                ADDRESS2: {}
            };

            expect(sheet.hyperlink("ADDRESS3", undefined)).toBe(sheet);
            expect(sheet._hyperlinks).toEqualJson({
                ADDRESS1: {},
                ADDRESS2: {}
            });

            sheet.hyperlink("ADDRESS1", undefined);
            expect(sheet._hyperlinks).toEqualJson({
                ADDRESS2: {}
            });

            sheet.hyperlink("ADDRESS2", undefined);
            expect(sheet._hyperlinks).toEqualJson({});
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

    describe("merged", () => {
        it("should return true/false if the cells are merged or not", () => {
            sheet._mergeCells = {
                ADDRESS1: {},
                ADDRESS2: {}
            };

            expect(sheet.merged("ADDRESS1")).toBe(true);
            expect(sheet.merged("ADDRESS2")).toBe(true);
            expect(sheet.merged("ADDRESS3")).toBe(false);
        });

        it("should add a mergeCell entry", () => {
            expect(sheet._mergeCells).toEqualJson({});
            expect(sheet.merged("ADDRESS", true)).toBe(sheet);
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

        it("should remove the mergeCell entry", () => {
            sheet._mergeCells = {
                ADDRESS1: {},
                ADDRESS2: {}
            };

            expect(sheet.merged("ADDRESS3", false)).toBe(sheet);
            expect(sheet._mergeCells).toEqualJson({
                ADDRESS1: {},
                ADDRESS2: {}
            });

            sheet.merged("ADDRESS1", false);
            expect(sheet._mergeCells).toEqualJson({
                ADDRESS2: {}
            });

            sheet.merged("ADDRESS2", false);
            expect(sheet._mergeCells).toEqualJson({});
        });
    });

    describe("toXmls", () => {
        it("should return the relationships", () => {
            expect(sheet.toXmls().relationships).toBe(sheet._relationships);
        });

        it("should return the ID node", () => {
            expect(sheet.toXmls().id).toBe(idNode);
        });

        it("should add the columns", () => {
            sheet._colsNode = {
                name: "cols",
                attributes: {},
                children: ["foo"]
            };
            sheet._colNodes = [
                null,
                {
                    name: "col",
                    attributes: { min: 1, max: 2, foo: true }
                },
                {
                    name: "col",
                    attributes: { min: 1, max: 2, foo: true }
                },
                {
                    name: "col",
                    attributes: { min: 3, max: 3, foo: true }
                },
                {
                    name: "col",
                    attributes: { min: 4, max: 4 }
                }
            ];

            expect(sheet.toXmls().sheet.children).toEqualJson([
                { name: 'sheetPr', attributes: {}, children: [] },
                { name: 'sheetFormatPr', attributes: {}, children: [] },
                {
                    name: 'cols',
                    attributes: {},
                    children: [
                        {
                            name: "col",
                            attributes: { min: 1, max: 2, foo: true }
                        },
                        {
                            name: "col",
                            attributes: { min: 3, max: 3, foo: true }
                        }
                    ]
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

            expect(sheet.toXmls().sheet.children).toEqualJson([
                { name: 'sheetPr', attributes: {}, children: [] },
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

        it("should add the hyperlinks", () => {
            sheet._hyperlinks = {
                A1: "HYPERLINK1",
                C3: "HYPERLINK2"
            };

            expect(sheet.toXmls().sheet.children).toEqualJson([
                { name: 'sheetPr', attributes: {}, children: [] },
                { name: 'sheetFormatPr', attributes: {}, children: [] },
                { name: 'sheetData', attributes: {}, children: [] },
                {
                    name: 'hyperlinks',
                    attributes: {},
                    children: ["HYPERLINK1", "HYPERLINK2"]
                },
                { name: 'pageMargins', attributes: {}, children: [] }
            ]);
        });

        it("should add the hyperlinks and merge cells in the proper order", () => {
            sheet._mergeCells = { "A1:B2": "MERGE1" };
            sheet._dataValidations.A1 = {
                name: "dataValidation",
                attributes: {
                    type: 'list',
                    sqref: 'A1'
                },
                children: ['STUFF']
            };
            sheet._hyperlinks = { A1: "HYPERLINK1" };

            expect(sheet.toXmls().sheet.children).toEqualJson([
                { name: 'sheetPr', attributes: {}, children: [] },
                { name: 'sheetFormatPr', attributes: {}, children: [] },
                { name: 'sheetData', attributes: {}, children: [] },
                {
                    name: 'mergeCells',
                    attributes: {},
                    children: ["MERGE1"]
                },
                {
                    name: 'dataValidations',
                    attributes: {},
                    children: [
                        {
                            name: "dataValidation",
                            attributes: {
                                type: 'list',
                                sqref: 'A1'
                            },
                            children: ['STUFF']
                        }
                    ]
                },
                {
                    name: 'hyperlinks',
                    attributes: {},
                    children: ["HYPERLINK1"]
                },
                { name: 'pageMargins', attributes: {}, children: [] }
            ]);
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

    describe("_getOrCreateSheetViewNode", () => {
        it("should get the existing sheet view node", () => {
            const sheetView = { name: "sheetView" };
            sheetNode.children.push({
                name: "sheetViews",
                attributes: {},
                children: [sheetView]
            });

            expect(sheet._getOrCreateSheetViewNode()).toBe(sheetView);
        });

        it("should create a new sheet view node", () => {
            const sheetView = sheet._getOrCreateSheetViewNode();
            expect(sheetView).toEqualJson({
                name: "sheetView",
                attributes: {
                    workbookViewId: 0
                },
                children: []
            });
            expect(sheetNode.children[1]).toEqualJson({
                name: "sheetViews",
                attributes: {},
                children: [sheetView]
            });
        });
    });

    describe("_init", () => {
        it("should create the sheet node", () => {
            sheet._init({}, {});
            expect(sheet._node).toEqual(jasmine.any(Object));
        });

        it("should create the relationships", () => {
            sheet._init({}, {}, {
                attributes: {},
                children: [
                    { name: "sheetData", attributes: {}, children: [] }
                ]
            }, "RELATIONSHIPS");

            expect(sheet._relationships).toEqual(jasmine.any(Relationships));
            expect(Relationships).toHaveBeenCalledWith("RELATIONSHIPS");
        });

        it("should delete the optional dimension node", () => {
            sheet._init({}, {}, {
                attributes: {},
                children: [
                    { name: "dimension", attributes: {}, children: [] },
                    { name: "sheetData", attributes: {}, children: [] }
                ]
            });

            expect(sheet._node.children).toEqualJson([
                { name: "sheetPr", attributes: {}, children: [] },
                { name: "sheetData", attributes: {}, children: [] }
            ]);
        });

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
            expect(sheet._sheetDataNode.children).toBe(sheet._rows);

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

            expect(sheet._colNodes).toEqualJson([
                undefined,
                undefined,
                { name: 'col', attributes: { min: 2, max: 3, foo: true } },
                { name: 'col', attributes: { min: 2, max: 3, foo: true } },
                undefined,
                { name: 'col', attributes: { min: 5, max: 5, bar: true } }
            ]);
        });

        it("should store the sheetPr node", () => {
            const sheetPrNode = { name: 'sheetPr', attributes: {}, children: [] };
            sheet._init({}, {}, {
                attributes: {},
                children: [
                    sheetPrNode,
                    { name: "sheetData", attributes: {}, children: [] }
                ]
            });
            
            expect(sheet._sheetPrNode).toBe(sheetPrNode);
        });

        it("should create the sheetPr node", () => {
            sheet._init({}, {}, {
                attributes: {},
                children: [
                    { name: "sheetData", attributes: {}, children: [] }
                ]
            });

            expect(sheet._sheetPrNode).toEqualJson({ name: 'sheetPr', attributes: {}, children: [] });
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

        it("should parse the hyperlinks", () => {
            sheet._init({}, {}, {
                attributes: {},
                children: [
                    { name: "sheetData", attributes: {}, children: [] },
                    {
                        name: 'hyperlinks',
                        attributes: {},
                        children: [
                            { name: 'hyperlink', attributes: { ref: "A1", foo: true } },
                            { name: 'hyperlink', attributes: { ref: "C3", bar: true } }
                        ]
                    }
                ]
            });

            expect(sheet._hyperlinks).toEqualJson({
                A1: { name: 'hyperlink', attributes: { ref: "A1", foo: true } },
                C3: { name: 'hyperlink', attributes: { ref: "C3", bar: true } }
            });
        });


        it("should parse the dataValidations", () => {
            sheet._init({}, {}, {
                attributes: {},
                children: [
                    { name: "sheetData", attributes: {}, children: [] },
                    {
                        name: 'dataValidations',
                        children: [
                            {
                                name: "dataValidation",
                                attributes: {
                                    type: 'list',
                                    allowBlank: false,
                                    showInputMessage: false,
                                    prompt: '',
                                    promptTitle: '',
                                    showErrorMessage: false,
                                    error: '',
                                    errorTitle: '',
                                    operator: '',
                                    sqref: 'A1'
                                },
                                children: [
                                    {
                                        name: 'formula1',
                                        atrributes: {},
                                        children: ['test1, test2, test3']
                                    },
                                    {
                                        name: 'formula2',
                                        atrributes: {},
                                        children: ['']
                                    }
                                ]
                            },
                            {
                                name: "dataValidation",
                                attributes: {
                                    type: 'list',
                                    allowBlank: false,
                                    showInputMessage: false,
                                    prompt: '',
                                    promptTitle: '',
                                    showErrorMessage: false,
                                    error: '',
                                    errorTitle: '',
                                    operator: '',
                                    sqref: 'A2'
                                },
                                children: [
                                    {
                                        name: 'formula1',
                                        atrributes: {},
                                        children: ['test1, test2, test3']
                                    },
                                    {
                                        name: 'formula2',
                                        atrributes: {},
                                        children: ['']
                                    }
                                ]
                            }
                        ]
                    }
                ]
            });

            expect(sheet._dataValidations).toEqualJson({
                A1: {
                    name: "dataValidation",
                    attributes: {
                        type: 'list',
                        allowBlank: false,
                        showInputMessage: false,
                        prompt: '',
                        promptTitle: '',
                        showErrorMessage: false,
                        error: '',
                        errorTitle: '',
                        operator: '',
                        sqref: 'A1'
                    },
                    children: [
                        {
                            name: 'formula1',
                            atrributes: {},
                            children: ['test1, test2, test3']
                        },
                        {
                            name: 'formula2',
                            atrributes: {},
                            children: ['']
                        }
                    ]
                },
                A2: {
                    name: "dataValidation",
                    attributes: {
                        type: 'list',
                        allowBlank: false,
                        showInputMessage: false,
                        prompt: '',
                        promptTitle: '',
                        showErrorMessage: false,
                        error: '',
                        errorTitle: '',
                        operator: '',
                        sqref: 'A2'
                    },
                    children: [
                        {
                            name: 'formula1',
                            atrributes: {},
                            children: ['test1, test2, test3']
                        },
                        {
                            name: 'formula2',
                            atrributes: {},
                            children: ['']
                        }
                    ]
                }
            });
        });
    });
});
