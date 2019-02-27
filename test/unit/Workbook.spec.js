"use strict";

const _ = require("lodash");
const proxyquire = require("proxyquire");
const Promise = require("jszip").external.Promise;

describe("Workbook", () => {
    let resolved, fs, externals, JSZip, workbookNode, Workbook, StyleSheet, Sheet, SharedStrings, Relationships, ContentTypes, CoreProperties, XmlParser, XmlBuilder, Encryptor, blank;

    beforeEach(() => {
        // Resolve a promise with a small random delay so they resolve out of order.
        resolved = val => {
            return new Promise(resolve => {
                setTimeout(resolve, Math.random() * 10);
            }).then(() => val);
        };

        JSZip = jasmine.createSpy("JSZip");
        JSZip.loadAsync = jasmine.createSpy("JSZip.loadAsync").and.returnValue(Promise.resolve(new JSZip()));
        JSZip.prototype.file = jasmine.createSpy("JSZip.file");
        JSZip.prototype.remove = jasmine.createSpy("JSZip.remove");
        JSZip.prototype.generateAsync = jasmine.createSpy("JSZip.generateAsync").and.returnValue(Promise.resolve("ZIP"));
        JSZip.external = { Promise };
        JSZip.prototype.file.and.callFake(fileName => ({
            async: () => Promise.resolve(`TEXT(${fileName})`)
        }));

        fs = jasmine.createSpyObj("fs", ["readFile", "writeFile"]);
        fs.readFile.and.callFake((path, cb) => cb(null, "DATA"));
        fs.writeFile.and.callFake((path, data, cb) => cb(null));

        StyleSheet = jasmine.createSpy("StyleSheet");
        StyleSheet.prototype.toString = () => "STYLE SHEET";

        Sheet = class {
            constructor(workbook, sheetIdNode, sheetNode, sheetRelationshipsNode) {
                this.workbook = workbook;
                this.sheetIdNode = sheetIdNode;
                this.sheetNode = sheetNode;
                this.sheetRelationshipsNode = sheetRelationshipsNode;
            }
        };
        Sheet.prototype.find = jasmine.createSpy("Sheet.find");
        let sheetOutput = false;
        Sheet.prototype.toXmls = jasmine.createSpy("Sheet.toXmls").and.callFake(() => {
            const relationships = sheetOutput ? "RELATIONSHIPS" : undefined;
            sheetOutput = !sheetOutput;
            return { sheet: "SHEET", id: { attributes: { 'r:id': "RID" } }, relationships };
        });
        Sheet.prototype.hidden = jasmine.createSpy("Sheet.hidden").and.returnValue(false);
        Sheet.prototype.tabSelected = jasmine.createSpy("Sheet.tabSelected");

        SharedStrings = jasmine.createSpy("SharedStrings");
        SharedStrings.prototype.toString = () => "SHARED STRINGS";

        Relationships = jasmine.createSpy("Relationships");
        Relationships.prototype.toString = () => "RELATIONSHIPS";
        Relationships.prototype.findByType = jasmine.createSpy("Relationships.findByType");
        Relationships.prototype.add = jasmine.createSpy("Relationships.add");

        ContentTypes = jasmine.createSpy("ContentTypes");
        ContentTypes.prototype.toString = () => "CONTENT TYPES";
        ContentTypes.prototype.findByPartName = jasmine.createSpy("ContentTypes.findByPartName");
        ContentTypes.prototype.add = jasmine.createSpy("ContentTypes.add");

        CoreProperties = jasmine.createSpy("CoreProperties");
        CoreProperties.prototype.toString = () => "CORE PROPERTIES";
        CoreProperties.prototype.get = jasmine.createSpy("CoreProperties.get");
        CoreProperties.prototype.set = jasmine.createSpy("CoreProperties.set");

        workbookNode = {
            name: 'workbook',
            attributes: {},
            children: [
                {
                    name: "bookViews",
                    attributes: {},
                    children: [
                        {
                            name: 'workbookView',
                            attributes: {},
                            children: []
                        }
                    ]
                },
                {
                    name: 'sheets',
                    attributes: {},
                    children: [
                        { name: 'sheet', attributes: { name: 'A', sheetId: 5 } },
                        { name: 'sheet', attributes: { name: 'B', sheetId: 9 } }
                    ]
                }
            ]
        };

        XmlParser = jasmine.createSpy("XmlParser");
        XmlParser.prototype.parseAsync = jasmine.createSpy("XmlParser.parseAsync").and.callFake(text => Promise.resolve(`JSON(${text})`));

        XmlBuilder = jasmine.createSpy("XmlBuilder");
        XmlBuilder.prototype.build = jasmine.createSpy("XmlBuilder.build").and.callFake(obj => `XML: ${obj && obj.toString()}`);

        Encryptor = jasmine.createSpy("Encryptor");
        Encryptor.prototype.encrypt = jasmine.createSpy("Encryptor.encrypt").and.callFake(input => `ENCRYPTED(${input})`);
        Encryptor.prototype.decryptAsync = jasmine.createSpy("Encryptor.decryptAsync").and.callFake(input => Promise.resolve(`DECRYPTED(${input})`));

        blank = () => "BLANK";

        // proxyquire doesn't like overriding raw objects... a spy obj works.
        externals = jasmine.createSpyObj("externals", ["_"]);
        externals.Promise = Promise;

        Workbook = proxyquire("../../lib/Workbook", {
            fs,
            jszip: JSZip,
            './externals': externals,
            './StyleSheet': StyleSheet,
            './Sheet': Sheet,
            './SharedStrings': SharedStrings,
            './Relationships': Relationships,
            './ContentTypes': ContentTypes,
            './CoreProperties': CoreProperties,
            './XmlParser': XmlParser,
            './XmlBuilder': XmlBuilder,
            './Encryptor': Encryptor,
            './blank': blank,
            '@noCallThru': true
        });
    });

    describe("static", () => {
        beforeEach(() => {
            spyOn(Workbook.prototype, "_initAsync").and.returnValue(Promise.resolve("WORKBOOK"));
        });

        describe("fromBlankAsync", () => {
            itAsync("should init with blank data", () => {
                return Workbook.fromBlankAsync()
                    .then(workbook => {
                        expect(Workbook.prototype._initAsync).toHaveBeenCalledWith("BLANK", undefined);
                        expect(workbook).toBe("WORKBOOK");
                    });
            });
        });

        describe("fromDataAsync", () => {
            itAsync("should init with the data", () => {
                return Workbook.fromDataAsync("DATA", "OPTS")
                    .then(workbook => {
                        expect(Workbook.prototype._initAsync).toHaveBeenCalledWith("DATA", "OPTS");
                        expect(workbook).toBe("WORKBOOK");
                    });
            });
        });

        describe("fromFileAsync", () => {
            if (process.browser) {
                it("should throw an error if in browser", () => {
                    expect(() => Workbook.fromFileAsync()).toThrow();
                });
            }

            if (!process.browser) {
                itAsync("should init with the file data", () => {
                    return Workbook.fromFileAsync("PATH", "OPTS")
                        .then(workbook => {
                            expect(Workbook.prototype._initAsync).toHaveBeenCalledWith("DATA", "OPTS");
                            expect(fs.readFile).toHaveBeenCalledWith("PATH", jasmine.any(Function));
                            expect(workbook).toBe("WORKBOOK");
                        });
                });
            }
        });
    });

    describe("instance", () => {
        let workbook;

        beforeEach(() => {
            workbook = new Workbook();
        });

        describe("activeSheet", () => {
            beforeEach(() => {
                workbook._node = workbookNode;
                workbook._sheets = [new Sheet(), new Sheet()];
                workbook._activeSheet = workbook._sheets[0];
            });

            it("should return the active sheet", () => {
                expect(workbook.activeSheet()).toBe(workbook._sheets[0]);
                workbook._activeSheet = workbook._sheets[1];
                expect(workbook.activeSheet()).toBe(workbook._sheets[1]);
            });

            it("should set the active sheet", () => {
                expect(workbook.activeSheet(workbook._sheets[1])).toBe(workbook);
                expect(workbook._sheets[0].tabSelected).toHaveBeenCalledWith(false);
                expect(workbook._sheets[1].tabSelected).toHaveBeenCalledWith(true);
                expect(workbook._activeSheet).toBe(workbook._sheets[1]);

                expect(workbook.activeSheet(0)).toBe(workbook);
                expect(workbook._activeSheet).toBe(workbook._sheets[0]);
            });
        });

        describe("addSheet", () => {
            beforeEach(() => {
                workbook._sheets = [new Sheet()];
                spyOn(workbook, "activeSheet").and.returnValue(workbook._sheets[0]);
                spyOn(workbook, "sheet");
                workbook._relationships = jasmine.createSpyObj("relationships", ["add"]);
                workbook._relationships.add.and.returnValue({
                    attributes: {
                        Id: 'RID'
                    }
                });
                workbook._maxSheetId = 7;
            });

            it("should throw an error if the sheet name is invalid", () => {
                expect(() => workbook.addSheet()).toThrow();
                expect(() => workbook.addSheet('foo?')).toThrow();
                expect(() => workbook.addSheet('12345678901234567890123456789012')).toThrow();

                expect(() => workbook.addSheet('foo')).not.toThrow();

                workbook.sheet.and.returnValue(workbook._sheets[0]);
                expect(() => workbook.addSheet('foo')).toThrow();
            });

            it("should add the sheet at the end", () => {
                const sheet = workbook.addSheet('foo');
                expect(sheet).toEqual(jasmine.any(Sheet));
                expect(workbook._sheets.length).toBe(2);
                expect(workbook._sheets[1]).toBe(sheet);
                expect(sheet.workbook).toBe(workbook);
                expect(sheet.sheetIdNode).toEqualJson({
                    name: "sheet",
                    attributes: {
                        name: "foo",
                        sheetId: 8,
                        'r:id': "RID"
                    },
                    children: []
                });
                expect(sheet.sheetNode).toBeUndefined();
                expect(sheet.sheetRelationshipsNode).toBeUndefined();
            });

            it("should add the sheet at the given index", () => {
                const sheet1 = workbook.addSheet('foo', 0);
                expect(workbook._sheets.length).toBe(2);
                expect(workbook._sheets[0]).toBe(sheet1);

                const sheet2 = workbook.addSheet('bar', 2);
                expect(workbook._sheets.length).toBe(3);
                expect(workbook._sheets[2]).toBe(sheet2);
            });

            it("should add the sheet before the given sheet", () => {
                const sheet = workbook.addSheet('foo', workbook._sheets[0]);
                expect(workbook._sheets.length).toBe(2);
                expect(workbook._sheets[0]).toBe(sheet);
            });

            it("should add the sheet before the sheet with the given name", () => {
                workbook.sheet.and.callFake(name => {
                    if (name === 'existing') return workbook._sheets[0];
                });

                const sheet = workbook.addSheet('foo', 'existing');
                expect(workbook._sheets.length).toBe(2);
                expect(workbook._sheets[0]).toBe(sheet);
            });
        });

        describe("definedName", () => {
            it("should return the scoped defined name", () => {
                spyOn(workbook, 'scopedDefinedName').and.returnValue("SCOPED DEFINED NAME");
                expect(workbook.definedName("NAME")).toBe("SCOPED DEFINED NAME");
                expect(workbook.scopedDefinedName).toHaveBeenCalledWith(undefined, "NAME");
            });
        });

        describe("deleteSheet", () => {
            let sheet1, sheet2, sheet3;
            beforeEach(() => {
                sheet1 = new Sheet();
                sheet2 = new Sheet();
                sheet3 = new Sheet();
                sheet1.name = jasmine.createSpy("name").and.returnValue("SHEET1");
                sheet2.name = jasmine.createSpy("name").and.returnValue("SHEET2");
                sheet3.name = jasmine.createSpy("name").and.returnValue("SHEET3");

                workbook._sheets = [sheet1, sheet2, sheet3];
                workbook._activeSheet = sheet2;
            });

            it("should throw an error if the sheet doesn't exist", () => {
                expect(() => workbook.deleteSheet("foo")).toThrow();
            });

            it("should throw an error if we are trying to hide the only visible sheet", () => {
                sheet1.hidden = jasmine.createSpy("hidden").and.returnValue(true);
                sheet2.hidden = jasmine.createSpy("hidden").and.returnValue(false);
                sheet3.hidden = jasmine.createSpy("hidden").and.returnValue(true);
                expect(() => workbook.deleteSheet(1)).toThrow();
                expect(() => workbook.deleteSheet(0)).not.toThrow();
            });

            it("should delete the sheet and update the active sheet as needed", () => {
                workbook.deleteSheet(1);
                expect(workbook._sheets).toEqual([sheet1, sheet3]);
                expect(workbook._activeSheet).toBe(sheet3);

                workbook.deleteSheet(0);
                expect(workbook._sheets).toEqual([sheet3]);
                expect(workbook._activeSheet).toBe(sheet3);
            });
        });

        describe("find", () => {
            it("should return the matches", () => {
                workbook._sheets = [
                    new Sheet(),
                    new Sheet(),
                    new Sheet()
                ];

                Sheet.prototype.find.and.returnValue(["A", "B"]);
                expect(workbook.find('foo')).toEqual(["A", "B", "A", "B", "A", "B"]);
                expect(Sheet.prototype.find).toHaveBeenCalledWith(/foo/gim, undefined);

                Sheet.prototype.find.and.returnValue('C');
                expect(workbook.find('bar', 'baz')).toEqual(['C', 'C', 'C']);
                expect(Sheet.prototype.find).toHaveBeenCalledWith(/bar/gim, 'baz');
            });
        });

        describe("moveSheet", () => {
            let sheet1, sheet2, sheet3;
            beforeEach(() => {
                sheet1 = new Sheet();
                sheet2 = new Sheet();
                sheet3 = new Sheet();
                sheet1.name = jasmine.createSpy("name").and.returnValue("SHEET1");
                sheet2.name = jasmine.createSpy("name").and.returnValue("SHEET2");
                sheet3.name = jasmine.createSpy("name").and.returnValue("SHEET3");

                workbook._sheets = [sheet1, sheet2, sheet3];
            });

            it("should throw an error if the sheet doesn't exist", () => {
                expect(() => workbook.moveSheet("foo")).toThrow();
                expect(() => workbook.moveSheet("SHEET1", "foo")).toThrow();
            });

            it("should move the sheet to the end", () => {
                workbook.moveSheet("SHEET2");
                expect(workbook._sheets).toEqual([sheet1, sheet3, sheet2]);
            });

            it("should move the sheet to the given index", () => {
                workbook.moveSheet("SHEET1", 1);
                expect(workbook._sheets).toEqual([sheet2, sheet1, sheet3]);
            });

            it("should move the sheet before the sheet with the given name", () => {
                workbook.moveSheet("SHEET3", "SHEET1");
                expect(workbook._sheets).toEqual([sheet3, sheet1, sheet2]);
            });

            it("should move the sheet before the given sheet", () => {
                workbook.moveSheet("SHEET2", sheet1);
                expect(workbook._sheets).toEqual([sheet2, sheet1, sheet3]);
            });
        });

        describe("outputAsync", () => {
            let relationships;

            beforeEach(() => {
                relationships =[];
                workbook._contentTypes = new ContentTypes();
                workbook._coreProperties = new CoreProperties();
                workbook._relationships = new Relationships();
                workbook._sharedStrings = new SharedStrings();
                workbook._styleSheet = new StyleSheet();
                workbook._node = "WORKBOOK";
                workbook._sheets = [new Sheet(), new Sheet()];
                workbook._sheetsNode = { name: 'sheets', attributes: {}, children: [] };
                workbook._zip = new JSZip();

                workbook._relationships.findById = jasmine.createSpy("findById").and.callFake(() => {
                    const relationship = { attributes: {} };
                    relationships.push(relationship);
                    return relationship;
                });
                spyOn(workbook, "_setSheetRefs");
                spyOn(workbook, "_convertBufferToOutput").and.returnValue("OUTPUT");
            });

            itAsync("should output the data", () => {
                return workbook.outputAsync({ type: "TYPE" })
                    .then(output => {
                        expect(output).toBe("OUTPUT");

                        expect(workbook._setSheetRefs).toHaveBeenCalledWith();

                        expect(workbook._zip.file).toHaveBeenCalledWith("[Content_Types].xml", "XML: CONTENT TYPES", { date: new Date(0), createFolders: false });
                        expect(workbook._zip.file).toHaveBeenCalledWith("docProps/core.xml", "XML: CORE PROPERTIES", { date: new Date(0), createFolders: false });
                        expect(workbook._zip.file).toHaveBeenCalledWith("xl/_rels/workbook.xml.rels", "XML: RELATIONSHIPS", { date: new Date(0), createFolders: false });
                        expect(workbook._zip.file).toHaveBeenCalledWith("xl/sharedStrings.xml", "XML: SHARED STRINGS", { date: new Date(0), createFolders: false });
                        expect(workbook._zip.file).toHaveBeenCalledWith("xl/styles.xml", "XML: STYLE SHEET", { date: new Date(0), createFolders: false });
                        expect(workbook._zip.file).toHaveBeenCalledWith("xl/workbook.xml", "XML: WORKBOOK", { date: new Date(0), createFolders: false });
                        expect(workbook._zip.file).toHaveBeenCalledWith("xl/worksheets/sheet1.xml", "XML: SHEET", { date: new Date(0), createFolders: false });
                        expect(workbook._zip.file).not.toHaveBeenCalledWith("xl/worksheets/_rels/sheet1.xml.rels", "XML: RELATIONSHIPS", { date: new Date(0), createFolders: false });
                        expect(workbook._zip.file).toHaveBeenCalledWith("xl/worksheets/sheet2.xml", "XML: SHEET", { date: new Date(0), createFolders: false });
                        expect(workbook._zip.file).toHaveBeenCalledWith("xl/worksheets/_rels/sheet2.xml.rels", "XML: RELATIONSHIPS", { date: new Date(0), createFolders: false });
                        expect(workbook._zip.generateAsync).toHaveBeenCalledWith({
                            type: 'nodebuffer',
                            compression: "DEFLATE"
                        });
                        expect(relationships).toEqualJson([
                            { attributes: { Target: "worksheets/sheet1.xml" } },
                            { attributes: { Target: "worksheets/sheet2.xml" } }
                        ]);
                        expect(workbook._sheetsNode.children).toEqualJson([
                            { attributes: { 'r:id': "RID" } },
                            { attributes: { 'r:id': "RID" } }
                        ]);
                        expect(Encryptor.prototype.encrypt).not.toHaveBeenCalled();
                        expect(workbook._convertBufferToOutput).toHaveBeenCalledWith("ZIP", "TYPE");
                    });
            });

            itAsync("should encrypt the workbook is password is set", () => {
                return workbook.outputAsync({ password: "PASSWORD" })
                    .then(output => {
                        expect(Encryptor.prototype.encrypt).toHaveBeenCalledWith("ZIP", "PASSWORD");
                        expect(workbook._convertBufferToOutput).toHaveBeenCalledWith("ENCRYPTED(ZIP)", undefined);
                        expect(output).toBe("OUTPUT");
                    });
            });
        });

        describe("sheet", () => {
            it("should return the matching sheet", () => {
                workbook._sheets = [{
                    name: () => "A"
                }, {
                    name: () => "B"
                }];

                expect(workbook.sheet(0)).toBe(workbook._sheets[0]);
                expect(workbook.sheet(1)).toBe(workbook._sheets[1]);
                expect(workbook.sheet("A")).toBe(workbook._sheets[0]);
                expect(workbook.sheet("B")).toBe(workbook._sheets[1]);
            });
        });

        describe("sheets", () => {
            it("should return the sheets", () => {
                workbook._sheets = ["SHEET1", "SHEET2"];
                expect(workbook.sheets()).toEqualJson(["SHEET1", "SHEET2"]);
                expect(workbook.sheets()).not.toBe(workbook._sheets);
            });
        });

        describe("toFileAsync", () => {
            if (process.browser) {
                it("should throw an error if in browser", () => {
                    expect(() => workbook.toFileAsync()).toThrow();
                });
            }

            if (!process.browser) {
                itAsync("should write the workbook to file", () => {
                    spyOn(workbook, "outputAsync").and.returnValue(Promise.resolve("OUTPUT"));
                    return workbook.toFileAsync("PATH")
                        .then(() => {
                            expect(fs.writeFile).toHaveBeenCalledWith("PATH", "OUTPUT", jasmine.any(Function));
                        });
                });
            }
        });

        describe("scopedDefinedName", () => {
            let sheet;

            beforeEach(() => {
                sheet = {
                    cell: jasmine.createSpy("cell").and.returnValue("CELL"),
                    range: jasmine.createSpy("range").and.returnValue("RANGE"),
                    row: jasmine.createSpy("row").and.returnValue("ROW"),
                    column: jasmine.createSpy("column").and.returnValue("COLUMN")
                };

                workbook._node = {
                    children: [{
                        name: 'definedNames',
                        children: [
                            { name: 'definedName', attributes: { name: 'cell' }, children: ["Sheet1!$A$1"] },
                            { name: 'definedName', attributes: { name: 'range' }, children: ["Sheet2!$A$1:B2"] },
                            { name: 'definedName', attributes: { name: 'column' }, children: ["Sheet3!$A:$A"] },
                            { name: 'definedName', attributes: { name: 'row' }, children: ["Sheet4!$1:$1"] },
                            { name: 'definedName', localSheet: sheet, attributes: { name: 'sheet scope' }, children: ["Sheet5!$A$1"] },
                            { name: 'definedName', attributes: { name: 'row range' }, children: ["Sheet1!$1:$3"] },
                            { name: 'definedName', attributes: { name: 'column range' }, children: ["Sheet1!$A:$C"] },
                            { name: 'definedName', attributes: { name: 'group' }, children: ["A1,A2"] },
                            { name: 'definedName', attributes: { name: 'formula' }, children: ["A1*A2"] }
                        ]
                    }]
                };

                spyOn(workbook, "sheet").and.returnValue(sheet);
            });

            it("should return undefined if not found", () => {
                expect(workbook.scopedDefinedName(undefined, "not found")).toBeUndefined();
            });

            it("should return the string if not supported", () => {
                expect(workbook.scopedDefinedName(undefined, "row range")).toEqual("Sheet1!$1:$3");
                expect(workbook.scopedDefinedName(undefined, "column range")).toEqual("Sheet1!$A:$C");
                expect(workbook.scopedDefinedName(undefined, "group")).toEqual("A1,A2");
                expect(workbook.scopedDefinedName(undefined, "formula")).toEqual("A1*A2");
            });

            it("should return the selection", () => {
                expect(workbook.scopedDefinedName(undefined, "cell")).toBe("CELL");
                expect(workbook.sheet).toHaveBeenCalledWith("Sheet1");
                expect(sheet.cell).toHaveBeenCalledWith(1, 1);

                expect(workbook.scopedDefinedName(undefined, "range")).toBe("RANGE");
                expect(workbook.sheet).toHaveBeenCalledWith("Sheet2");
                expect(sheet.range).toHaveBeenCalledWith(1, 1, 2, 2);

                expect(workbook.scopedDefinedName(undefined, "column")).toBe("COLUMN");
                expect(workbook.sheet).toHaveBeenCalledWith("Sheet3");
                expect(sheet.column).toHaveBeenCalledWith(1);

                expect(workbook.scopedDefinedName(undefined, "row")).toBe("ROW");
                expect(workbook.sheet).toHaveBeenCalledWith("Sheet3");
                expect(sheet.row).toHaveBeenCalledWith(1);
            });

            it("should return the scoped selection", () => {
                expect(workbook.scopedDefinedName(undefined, "sheet scope")).toBeUndefined();

                expect(workbook.scopedDefinedName({}, "sheet scope")).toBeUndefined();

                expect(workbook.scopedDefinedName(sheet, "sheet scope")).toBe("CELL");
            });

            it("should set the defined name with a string", () => {
                expect(workbook.scopedDefinedName(undefined, "NAME", "VALUE")).toBe(workbook);
                expect(workbook._node.children[0].children[9]).toEqualJson({
                    name: "definedName",
                    attributes: { name: "NAME" },
                    children: ["VALUE"]
                });
            });

            it("should define a sheet scoped name", () => {
                expect(workbook.scopedDefinedName(sheet, "NAME", "VALUE")).toBe(workbook);
                expect(workbook._node.children[0].children[9]).toEqualJson({
                    name: "definedName",
                    attributes: { name: "NAME" },
                    children: ["VALUE"],
                    localSheet: {}
                });
                expect(workbook._node.children[0].children[9].localSheet).toBe(sheet);
            });

            it("should set the defined name with a cell", () => {
                const cell = jasmine.createSpyObj("cell", ["address"]);
                cell.address.and.returnValue("ADDRESS");

                expect(workbook.scopedDefinedName(undefined, "NAME", cell)).toBe(workbook);
                expect(workbook._node.children[0].children[9]).toEqualJson({
                    name: "definedName",
                    attributes: { name: "NAME" },
                    children: ["ADDRESS"]
                });
                expect(cell.address).toHaveBeenCalledWith({ includeSheetName: true, anchored: true });
            });

            it("should unset a name", () => {
                workbook._node.children[0].children.length = 2;

                workbook.scopedDefinedName(undefined, "cell", null);
                expect(workbook._node.children[0].children.length).toBe(1);

                workbook.scopedDefinedName(undefined, "range", null);
                expect(workbook._node.children.length).toBe(0);
            });
        });

        describe("sharedStrings", () => {
            it("should return the shared strings", () => {
                workbook._sharedStrings = "SHARED STRINGS";
                expect(workbook.sharedStrings()).toBe("SHARED STRINGS");
            });
        });

        describe("styleSheet", () => {
            it("should return the style sheet", () => {
                workbook._styleSheet = "STYLE SHEET";
                expect(workbook.styleSheet()).toBe("STYLE SHEET");
            });
        });

        describe("coreProperties", () => {
            it("should return the core properties", () => {
                workbook._coreProperties = "CORE PROPERTIES";
                expect(workbook.properties()).toBe("CORE PROPERTIES");
            });
        });

        describe("_initAsync", () => {
            beforeEach(() => {
                spyOn(workbook, "_parseNodesAsync").and.callFake(files => {
                    return Promise.all(_.map(files, file => {
                        if (file === "xl/workbook.xml") return resolved(workbookNode);
                        return resolved(`PARSED(${file})`);
                    }));
                });
                spyOn(workbook, "_parseSheetRefs");
                spyOn(workbook, "_convertInputToBufferAsync").and.returnValue(Promise.resolve("BUFFER"));
            });

            itAsync("should extract the files from the data zip and load the objects", () => {
                return workbook._initAsync("DATA", { base64: "BASE64" })
                    .then(wb => {
                        expect(wb).toBe(workbook);

                        expect(workbook._convertInputToBufferAsync).toHaveBeenCalledWith("DATA", "BASE64");

                        expect(Encryptor.prototype.decryptAsync).not.toHaveBeenCalled();

                        expect(JSZip.loadAsync).toHaveBeenCalledWith("BUFFER");
                        expect(workbook._zip).toEqual(jasmine.any(JSZip));

                        expect(workbook._contentTypes).toEqual(jasmine.any(ContentTypes));
                        expect(workbook._relationships).toEqual(jasmine.any(Relationships));
                        expect(workbook._sharedStrings).toEqual(jasmine.any(SharedStrings));
                        expect(workbook._styleSheet).toEqual(jasmine.any(StyleSheet));
                        expect(workbook._sheets[0]).toEqual(jasmine.any(Sheet));
                        expect(workbook._sheets[1]).toEqual(jasmine.any(Sheet));
                        expect(workbook._node).toBe(workbookNode);

                        expect(workbook._sheets[0].workbook).toBe(workbook);
                        expect(workbook._sheets[0].sheetIdNode).toEqual({ name: 'sheet', attributes: { name: 'A', sheetId: 5 } });
                        expect(workbook._sheets[0].sheetNode).toEqual('PARSED(xl/worksheets/sheet1.xml)');
                        expect(workbook._sheets[0].sheetRelationshipsNode).toEqual('PARSED(xl/worksheets/_rels/sheet1.xml.rels)');
                        expect(workbook._sheets[1].workbook).toBe(workbook);
                        expect(workbook._sheets[1].sheetIdNode).toEqual({ name: 'sheet', attributes: { name: 'B', sheetId: 9 } });
                        expect(workbook._sheets[1].sheetNode).toEqual('PARSED(xl/worksheets/sheet2.xml)');
                        expect(workbook._sheets[1].sheetRelationshipsNode).toEqual('PARSED(xl/worksheets/_rels/sheet2.xml.rels)');

                        expect(ContentTypes).toHaveBeenCalledWith('PARSED([Content_Types].xml)');
                        expect(Relationships).toHaveBeenCalledWith('PARSED(xl/_rels/workbook.xml.rels)');
                        expect(SharedStrings).toHaveBeenCalledWith('PARSED(xl/sharedStrings.xml)');
                        expect(StyleSheet).toHaveBeenCalledWith('PARSED(xl/styles.xml)');

                        expect(Relationships.prototype.findByType).toHaveBeenCalledWith('sharedStrings');
                        expect(ContentTypes.prototype.findByPartName).toHaveBeenCalledWith("/xl/sharedStrings.xml");

                        expect(workbook._zip.remove).toHaveBeenCalledWith("xl/calcChain.xml");

                        expect(workbook._maxSheetId).toBe(9);

                        expect(workbook._parseSheetRefs).toHaveBeenCalledWith();
                    });
            });

            itAsync("should decrypte the data if a password is set", () => {
                return workbook._initAsync("DATA", { password: "PASSWORD" })
                    .then(wb => {
                        expect(wb).toBe(workbook);
                        expect(workbook._convertInputToBufferAsync).toHaveBeenCalledWith("DATA", undefined);
                        expect(Encryptor.prototype.decryptAsync).toHaveBeenCalledWith("BUFFER", "PASSWORD");
                        expect(JSZip.loadAsync).toHaveBeenCalledWith("DECRYPTED(BUFFER)");
                    });
            });

            itAsync("should not add the shared strings if already present", () => {
                Relationships.prototype.findByType.and.returnValue({});
                ContentTypes.prototype.findByPartName.and.returnValue({});

                return workbook._initAsync("DATA")
                    .then(() => {
                        expect(Relationships.prototype.add).not.toHaveBeenCalled();
                        expect(ContentTypes.prototype.add).not.toHaveBeenCalled();
                    });
            });

            itAsync("should not add the shared strings if already present", () => {
                Relationships.prototype.findByType.and.returnValue(undefined);
                ContentTypes.prototype.findByPartName.and.returnValue(undefined);

                return workbook._initAsync("DATA")
                    .then(() => {
                        expect(Relationships.prototype.add).toHaveBeenCalledWith("sharedStrings", "sharedStrings.xml");
                        expect(ContentTypes.prototype.add).toHaveBeenCalledWith("/xl/sharedStrings.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
                    });
            });
        });

        describe("_parseNodesAsync", () => {
            itAsync("should parse the nodes", () => {
                workbook._zip = new JSZip();
                return workbook._parseNodesAsync(["foo", "bar", "baz"])
                    .then(nodes => {
                        expect(workbook._zip.file).toHaveBeenCalledWith("foo");
                        expect(workbook._zip.file).toHaveBeenCalledWith("bar");
                        expect(workbook._zip.file).toHaveBeenCalledWith("baz");

                        expect(XmlParser.prototype.parseAsync).toHaveBeenCalledWith('TEXT(foo)');
                        expect(XmlParser.prototype.parseAsync).toHaveBeenCalledWith('TEXT(bar)');
                        expect(XmlParser.prototype.parseAsync).toHaveBeenCalledWith('TEXT(baz)');

                        expect(nodes).toEqualJson([
                            "JSON(TEXT(foo))",
                            "JSON(TEXT(bar))",
                            "JSON(TEXT(baz))"
                        ]);
                    });
            });
        });

        describe("_parseSheetRefs", () => {
            beforeEach(() => {
                workbook._node = {
                    name: 'workbook',
                    attributes: {},
                    children: []
                };

                workbook._sheets = ["SHEET1", "SHEET2"];
            });

            it("should parse the active sheet", () => {
                workbook._node.children = [{
                    name: "bookViews",
                    attributes: {},
                    children: [{
                        name: "workbookView",
                        attributes: { activeTab: 1 },
                        children: []
                    }]
                }];

                workbook._parseSheetRefs();
                expect(workbook._activeSheet).toBe("SHEET2");
            });

            it("should parse the defined names sheets", () => {
                workbook._node.children = [{
                    name: "definedNames",
                    attributes: {},
                    children: [
                        {
                            name: "definedName",
                            attributes: { name: "WORKBOOK_SCOPE" },
                            children: ["VALUE1"]
                        },
                        {
                            name: "definedName",
                            attributes: { name: "SHEET_SCOPE", localSheetId: 0 },
                            children: ["VALUE2"]
                        }
                    ]
                }];

                workbook._parseSheetRefs();

                expect(workbook._node.children).toEqualJson([{
                    name: "definedNames",
                    attributes: {},
                    children: [
                        {
                            name: "definedName",
                            attributes: { name: "WORKBOOK_SCOPE" },
                            children: ["VALUE1"]
                        },
                        {
                            name: "definedName",
                            attributes: { name: "SHEET_SCOPE", localSheetId: 0 },
                            children: ["VALUE2"],
                            localSheet: "SHEET1"
                        }
                    ]
                }]);
            });
        });

        describe("_setSheetRefs", () => {
            beforeEach(() => {
                workbook._node = {
                    name: 'workbook',
                    attributes: {},
                    children: []
                };

                workbook._sheets = ["SHEET1", "SHEET2"];
                workbook._activeSheet = "SHEET2";
            });

            it("should set the active sheet and create the book view", () => {
                workbook._setSheetRefs();
                expect(workbook._node.children).toEqualJson([{
                    name: "bookViews",
                    attributes: {},
                    children: [{
                        name: "workbookView",
                        attributes: { activeTab: 1 },
                        children: []
                    }]
                }]);
            });

            it("should set the active sheet and use the existing book view", () => {
                workbook._node.children = [{
                    name: "bookViews",
                    attributes: { foo: true },
                    children: []
                }];

                workbook._setSheetRefs();
                expect(workbook._node.children).toEqualJson([{
                    name: "bookViews",
                    attributes: { foo: true },
                    children: [{
                        name: "workbookView",
                        attributes: { activeTab: 1 },
                        children: []
                    }]
                }]);
            });

            it("should set the defined names sheets", () => {
                workbook._node.children = [{
                    name: "definedNames",
                    attributes: {},
                    children: [
                        {
                            name: "definedName",
                            attributes: { name: "WORKBOOK_SCOPE" },
                            children: ["VALUE1"]
                        },
                        {
                            name: "definedName",
                            attributes: { name: "SHEET_SCOPE" },
                            children: ["VALUE2"],
                            localSheet: "SHEET1"
                        }
                    ]
                }];

                workbook._setSheetRefs();

                expect(workbook._node.children[1]).toEqualJson({
                    name: "definedNames",
                    attributes: {},
                    children: [
                        {
                            name: "definedName",
                            attributes: { name: "WORKBOOK_SCOPE" },
                            children: ["VALUE1"]
                        },
                        {
                            name: "definedName",
                            attributes: { name: "SHEET_SCOPE", localSheetId: 0 },
                            children: ["VALUE2"],
                            localSheet: "SHEET1"
                        }
                    ]
                });
            });
        });

        describe("_convertBufferToOutput", () => {
            if (!process.browser) {
                it("should default to buffer in Node", () => {
                    const input = Buffer.alloc(5);
                    const output = workbook._convertBufferToOutput(input);
                    expect(Buffer.isBuffer(output)).toBe(true);
                });
            }

            if (process.browser) {
                it("should default to blob in browser", () => {
                    const input = Buffer.alloc(5);
                    const output = workbook._convertBufferToOutput(input);
                    expect(output).toEqual(jasmine.any(Blob));
                });
            }

            it("should return buffers unchanged", () => {
                const input = Buffer.alloc(5);
                const output = workbook._convertBufferToOutput(input, "nodebuffer");
                expect(output).toBe(input);
            });

            if (process.browser) {
                itAsync("should convert to a blob", () => {
                    const input = Buffer.from([0x66, 0x6f, 0x6f, 0x62, 0x61, 0x72]);
                    const output = workbook._convertBufferToOutput(input, "blob");
                    expect(output).toEqual(jasmine.any(Blob));

                    return new Promise(resolve => {
                        const fileReader = new FileReader();
                        fileReader.onload = event => {
                            resolve(event.target.result);
                        };
                        fileReader.readAsArrayBuffer(output);
                    }).then(buffer => {
                        expect(new Uint8Array(buffer)).toEqualUInt8Array(input);
                    });
                });
            }

            it("should convert to a base64 string", () => {
                const input = Buffer.from([0x66, 0x6f, 0x6f, 0x62, 0x61, 0x72]);
                const output = workbook._convertBufferToOutput(input, "base64");
                expect(output).toEqual("Zm9vYmFy");
            });

            it("should convert to a binary string", () => {
                const input = Buffer.from([0x66, 0x6f, 0x6f, 0x62, 0x61, 0x72]);
                const output = workbook._convertBufferToOutput(input, "binarystring");
                expect(output).toEqual("foobar");
            });

            it("should convert to a Uint8Array", () => {
                const input = Buffer.from([0x66, 0x6f, 0x6f, 0x62, 0x61, 0x72]);
                const output = workbook._convertBufferToOutput(input, "uint8array");
                expect(output).toEqual(jasmine.any(Uint8Array));
                expect(output).toEqualUInt8Array(input);
            });

            it("should convert to an ArrayBuffer", () => {
                const input = Buffer.from([0x66, 0x6f, 0x6f, 0x62, 0x61, 0x72]);
                const output = workbook._convertBufferToOutput(input, "arraybuffer");
                expect(output).toEqual(jasmine.any(ArrayBuffer));
                expect(new Uint8Array(output)).toEqualUInt8Array(input);
            });
        });

        describe("_convertInputToBufferAsync", () => {
            itAsync("should return buffers unchanged", () => {
                const input = Buffer.alloc(5);
                return workbook._convertInputToBufferAsync(input)
                    .then(output => {
                        expect(output).toBe(input);
                    });
            });

            if (process.browser) {
                itAsync("should convert a blob", () => {
                    const input = new Blob([new Uint8Array([0x66, 0x6f, 0x6f, 0x62, 0x61, 0x72])], { type: Workbook.MIME_TYPE });
                    return workbook._convertInputToBufferAsync(input)
                        .then(output => {
                            expect(Buffer.isBuffer(output)).toBe(true);
                            expect(output).toEqualUInt8Array(Buffer.from([0x66, 0x6f, 0x6f, 0x62, 0x61, 0x72]));
                        });
                });
            }

            itAsync("should convert a base64 string", () => {
                return workbook._convertInputToBufferAsync("Zm9vYmFy", true)
                    .then(output => {
                        expect(Buffer.isBuffer(output)).toBe(true);
                        expect(output).toEqualUInt8Array(Buffer.from([0x66, 0x6f, 0x6f, 0x62, 0x61, 0x72]));
                    });
            });

            itAsync("should convert a binary string", () => {
                return workbook._convertInputToBufferAsync("foobar")
                    .then(output => {
                        expect(Buffer.isBuffer(output)).toBe(true);
                        expect(output).toEqualUInt8Array(Buffer.from([0x66, 0x6f, 0x6f, 0x62, 0x61, 0x72]));
                    });
            });

            itAsync("should convert a Uint8Array", () => {
                const input = new Uint8Array([0x66, 0x6f, 0x6f, 0x62, 0x61, 0x72])
                return workbook._convertInputToBufferAsync(input)
                    .then(output => {
                        expect(Buffer.isBuffer(output)).toBe(true);
                        expect(output).toEqualUInt8Array(Buffer.from([0x66, 0x6f, 0x6f, 0x62, 0x61, 0x72]));
                    });
            });

            itAsync("should convert an ArrayBuffer", () => {
                const input = new Uint8Array([0x66, 0x6f, 0x6f, 0x62, 0x61, 0x72]).buffer;
                return workbook._convertInputToBufferAsync(input)
                    .then(output => {
                        expect(Buffer.isBuffer(output)).toBe(true);
                        expect(output).toEqualUInt8Array(Buffer.from([0x66, 0x6f, 0x6f, 0x62, 0x61, 0x72]));
                    });
            });
        });

        describe('cloneSheet', () => {
            beforeEach(() => {
                workbook._sheets = [new Sheet()];
                spyOn(workbook, "activeSheet").and.returnValue(workbook._sheets[0]);
                spyOn(workbook, "sheet");
                workbook._relationships = jasmine.createSpyObj("relationships", ["add"]);
                workbook._relationships.add.and.returnValue({
                    attributes: {
                        Id: 'RID'
                    }
                });
            });

            it("should throw an error if params are invalid", () => {
                expect(() => workbook.cloneSheet()).toThrow();
                const from = workbook.addSheet('foo');
                expect(() => workbook.cloneSheet(from)).toThrow();
            });

            it("should add the sheet at the end", () => {
                const from = workbook._sheets[0];
                const sheet = workbook.cloneSheet(from, 'foo');
                expect(sheet).toEqual(jasmine.any(Sheet));
                expect(workbook._sheets.length).toBe(2);
                expect(workbook._sheets[1]).toBe(sheet);
                expect(sheet.workbook).toBe(workbook);
            });

            it("should add the sheet before the given sheet", () => {
                const from = workbook._sheets[0];
                const sheet = workbook.cloneSheet(from, 'foo', workbook._sheets[0]);
                expect(workbook._sheets.length).toBe(2);
                expect(workbook._sheets[0]).toBe(sheet);
            });

        });
    });
});
