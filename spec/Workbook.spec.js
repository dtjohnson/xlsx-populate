"use strict";

const _ = require("lodash");
const proxyquire = require("proxyquire");
const Promise = require("jszip").external.Promise;

describe("Workbook", () => {
    let resolved, fs, externals, JSZip, workbookNode, Workbook, StyleSheet, Sheet, SharedStrings, Relationships, ContentTypes, XmlParser, XmlBuilder, blank;

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
        StyleSheet.prototype.toObject = jasmine.createSpy("StyleSheet.toObject").and.returnValue("STYLE SHEET");

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
        Sheet.prototype.toObject = jasmine.createSpy("Sheet.toObject").and.callFake(() => {
            const relationships = sheetOutput ? "RELATIONSHIPS" : undefined;
            sheetOutput = !sheetOutput;
            return { sheet: "SHEET", relationships };
        });

        SharedStrings = jasmine.createSpy("SharedStrings");
        SharedStrings.prototype.toObject = jasmine.createSpy("SharedStrings.toObject").and.returnValue("SHARED STRINGS");

        Relationships = jasmine.createSpy("Relationships");
        Relationships.prototype.toObject = jasmine.createSpy("Relationships.toObject").and.returnValue("RELATIONSHIPS");
        Relationships.prototype.findByType = jasmine.createSpy("Relationships.findByType");
        Relationships.prototype.add = jasmine.createSpy("Relationships.add");

        ContentTypes = jasmine.createSpy("ContentTypes");
        ContentTypes.prototype.toObject = jasmine.createSpy("ContentTypes.toObject").and.returnValue("CONTENT TYPES");
        ContentTypes.prototype.findByPartName = jasmine.createSpy("ContentTypes.findByPartName");
        ContentTypes.prototype.add = jasmine.createSpy("ContentTypes.add");

        workbookNode = {
            name: 'workbook',
            attributes: {},
            children: [{
                name: 'sheets',
                attributes: {},
                children: [
                    { name: 'sheet', attributes: { name: 'A' } },
                    { name: 'sheet', attributes: { name: 'B' } }
                ]
            }]
        };

        XmlParser = jasmine.createSpy("XmlParser");
        XmlParser.prototype.parseAsync = jasmine.createSpy("XmlParser.parseAsync").and.callFake(text => Promise.resolve(`JSON(${text})`));

        XmlBuilder = jasmine.createSpy("XmlBuilder");
        XmlBuilder.prototype.build = jasmine.createSpy("XmlBuilder.build").and.callFake(obj => `XML: ${obj}`);

        blank = () => "BLANK";

        // proxyquire doesn't like overriding raw objects... a spy obj works.
        externals = jasmine.createSpyObj("externals", ["_"]);
        externals.Promise = Promise;

        Workbook = proxyquire("../lib/Workbook", {
            fs,
            jszip: JSZip,
            './externals': externals,
            './StyleSheet': StyleSheet,
            './Sheet': Sheet,
            './SharedStrings': SharedStrings,
            './Relationships': Relationships,
            './ContentTypes': ContentTypes,
            './XmlParser': XmlParser,
            './XmlBuilder': XmlBuilder,
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
                        expect(Workbook.prototype._initAsync).toHaveBeenCalledWith("BLANK");
                        expect(workbook).toBe("WORKBOOK");
                    });
            });
        });

        describe("fromDataAsync", () => {
            itAsync("should init with the data", () => {
                return Workbook.fromDataAsync("DATA")
                    .then(workbook => {
                        expect(Workbook.prototype._initAsync).toHaveBeenCalledWith("DATA");
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
                    return Workbook.fromFileAsync("PATH")
                        .then(workbook => {
                            expect(Workbook.prototype._initAsync).toHaveBeenCalledWith("DATA");
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

        describe("definedName", () => {
            it("should return the scoped defined name", () => {
                spyOn(workbook, 'scopedDefinedName').and.returnValue("SCOPED DEFINED NAME");
                expect(workbook.definedName("NAME")).toBe("SCOPED DEFINED NAME");
                expect(workbook.scopedDefinedName).toHaveBeenCalledWith("NAME");
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

        describe("outputAsync", () => {
            beforeEach(() => {
                workbook._contentTypes = new ContentTypes();
                workbook._relationships = new Relationships();
                workbook._sharedStrings = new SharedStrings();
                workbook._styleSheet = new StyleSheet();
                workbook._node = "WORKBOOK";
                workbook._sheets = [new Sheet(), new Sheet()];
                workbook._zip = new JSZip();
            });

            itAsync("should output the data", () => {
                return workbook.outputAsync('TYPE')
                    .then(() => {
                        expect(workbook._zip.file).toHaveBeenCalledWith("[Content_Types].xml", "XML: CONTENT TYPES", { date: new Date(0), createFolders: false });
                        expect(workbook._zip.file).toHaveBeenCalledWith("xl/_rels/workbook.xml.rels", "XML: RELATIONSHIPS", { date: new Date(0), createFolders: false });
                        expect(workbook._zip.file).toHaveBeenCalledWith("xl/sharedStrings.xml", "XML: SHARED STRINGS", { date: new Date(0), createFolders: false });
                        expect(workbook._zip.file).toHaveBeenCalledWith("xl/styles.xml", "XML: STYLE SHEET", { date: new Date(0), createFolders: false });
                        expect(workbook._zip.file).toHaveBeenCalledWith("xl/workbook.xml", "XML: WORKBOOK", { date: new Date(0), createFolders: false });
                        expect(workbook._zip.file).toHaveBeenCalledWith("xl/worksheets/sheet1.xml", "XML: SHEET", { date: new Date(0), createFolders: false });
                        expect(workbook._zip.file).not.toHaveBeenCalledWith("xl/worksheets/_rels/sheet1.xml.rels", "XML: RELATIONSHIPS", { date: new Date(0), createFolders: false });
                        expect(workbook._zip.file).toHaveBeenCalledWith("xl/worksheets/sheet2.xml", "XML: SHEET", { date: new Date(0), createFolders: false });
                        expect(workbook._zip.file).toHaveBeenCalledWith("xl/worksheets/_rels/sheet2.xml.rels", "XML: RELATIONSHIPS", { date: new Date(0), createFolders: false });
                        expect(workbook._zip.generateAsync).toHaveBeenCalledWith({
                            type: 'TYPE',
                            compression: "DEFLATE",
                            mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        });
                    });
            });

            if (!process.browser) {
                itAsync("should default type to buffer if node", () => {
                    return workbook.outputAsync()
                        .then(() => {
                            expect(workbook._zip.generateAsync).toHaveBeenCalledWith({
                                type: 'nodebuffer',
                                compression: "DEFLATE",
                                mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            });
                        });
                });
            }

            if (process.browser) {
                itAsync("should default type to blob if browser", () => {
                    return workbook.outputAsync()
                        .then(() => {
                            expect(workbook._zip.generateAsync).toHaveBeenCalledWith({
                                type: 'blob',
                                compression: "DEFLATE",
                                mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            });
                        });
                });
            }
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
                workbook._node = {
                    children: [{
                        name: 'definedNames',
                        children: [
                            { name: 'definedName', attributes: { name: 'cell' }, children: ["Sheet1!$A$1"] },
                            { name: 'definedName', attributes: { name: 'range' }, children: ["Sheet2!$A$1:B2"] },
                            { name: 'definedName', attributes: { name: 'column' }, children: ["Sheet3!$A:$A"] },
                            { name: 'definedName', attributes: { name: 'row' }, children: ["Sheet4!$1:$1"] },
                            { name: 'definedName', attributes: { name: 'sheet scope', localSheetId: 2 }, children: ["Sheet5!$A$1"] },
                            { name: 'definedName', attributes: { name: 'row range' }, children: ["Sheet1!$1:$3"] },
                            { name: 'definedName', attributes: { name: 'column range' }, children: ["Sheet1!$A:$C"] },
                            { name: 'definedName', attributes: { name: 'group' }, children: ["A1,A2"] },
                            { name: 'definedName', attributes: { name: 'formula' }, children: ["A1*A2"] }
                        ]
                    }]
                };

                sheet = {
                    cell: jasmine.createSpy("cell").and.returnValue("CELL"),
                    range: jasmine.createSpy("range").and.returnValue("RANGE"),
                    row: jasmine.createSpy("row").and.returnValue("ROW"),
                    column: jasmine.createSpy("column").and.returnValue("COLUMN"),
                };
                spyOn(workbook, "sheet").and.returnValue(sheet);
            });

            it("should return undefined if not found", () => {
                expect(workbook.scopedDefinedName("not found")).toBeUndefined();
            });

            it("should throw an error if not supported", () => {
                expect(() => workbook.scopedDefinedName("row range")).toThrow();
                expect(() => workbook.scopedDefinedName("column range")).toThrow();
                expect(() => workbook.scopedDefinedName("group")).toThrow();
                expect(() => workbook.scopedDefinedName("formula")).toThrow();
            });

            it("should return the selection", () => {
                expect(workbook.scopedDefinedName("cell")).toBe("CELL");
                expect(workbook.sheet).toHaveBeenCalledWith("Sheet1");
                expect(sheet.cell).toHaveBeenCalledWith(1, 1);

                expect(workbook.scopedDefinedName("range")).toBe("RANGE");
                expect(workbook.sheet).toHaveBeenCalledWith("Sheet2");
                expect(sheet.range).toHaveBeenCalledWith(1, 1, 2, 2);

                expect(workbook.scopedDefinedName("column")).toBe("COLUMN");
                expect(workbook.sheet).toHaveBeenCalledWith("Sheet3");
                expect(sheet.column).toHaveBeenCalledWith(1);

                expect(workbook.scopedDefinedName("row")).toBe("ROW");
                expect(workbook.sheet).toHaveBeenCalledWith("Sheet3");
                expect(sheet.row).toHaveBeenCalledWith(1);
            });

            it("should return the scoped selection", () => {
                expect(workbook.scopedDefinedName("sheet scope")).toBeUndefined();

                workbook._sheets = [undefined, sheet];
                expect(workbook.scopedDefinedName("sheet scope", sheet)).toBeUndefined();

                workbook._sheets = [undefined, undefined, sheet];
                expect(workbook.scopedDefinedName("sheet scope", sheet)).toBe("CELL");
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

        describe("_initAsync", () => {
            beforeEach(() => {
                spyOn(workbook, "_parseNodesAsync").and.callFake(files => {
                    return Promise.all(_.map(files, file => {
                        if (file === "xl/workbook.xml") return resolved(workbookNode);
                        return resolved(`PARSED(${file})`);
                    }));
                });
            });

            itAsync("should extract the files from the data zip and load the objects", () => {
                return workbook._initAsync("DATA")
                    .then(() => {
                        expect(JSZip.loadAsync).toHaveBeenCalledWith("DATA");
                        expect(workbook._zip).toEqual(jasmine.any(JSZip));

                        expect(workbook._contentTypes).toEqual(jasmine.any(ContentTypes));
                        expect(workbook._relationships).toEqual(jasmine.any(Relationships));
                        expect(workbook._sharedStrings).toEqual(jasmine.any(SharedStrings));
                        expect(workbook._styleSheet).toEqual(jasmine.any(StyleSheet));
                        expect(workbook._sheets[0]).toEqual(jasmine.any(Sheet));
                        expect(workbook._sheets[1]).toEqual(jasmine.any(Sheet));
                        expect(workbook._node).toBe(workbookNode);

                        expect(workbook._sheets[0].workbook).toBe(workbook);
                        expect(workbook._sheets[0].sheetIdNode).toEqual({ name: 'sheet', attributes: { name: 'A' } });
                        expect(workbook._sheets[0].sheetNode).toEqual('PARSED(xl/worksheets/sheet1.xml)');
                        expect(workbook._sheets[0].sheetRelationshipsNode).toEqual('PARSED(xl/worksheets/_rels/sheet1.xml.rels)');
                        expect(workbook._sheets[1].workbook).toBe(workbook);
                        expect(workbook._sheets[1].sheetIdNode).toEqual({ name: 'sheet', attributes: { name: 'B' } });
                        expect(workbook._sheets[1].sheetNode).toEqual('PARSED(xl/worksheets/sheet2.xml)');
                        expect(workbook._sheets[1].sheetRelationshipsNode).toEqual('PARSED(xl/worksheets/_rels/sheet2.xml.rels)');

                        expect(ContentTypes).toHaveBeenCalledWith('PARSED([Content_Types].xml)');
                        expect(Relationships).toHaveBeenCalledWith('PARSED(xl/_rels/workbook.xml.rels)');
                        expect(SharedStrings).toHaveBeenCalledWith('PARSED(xl/sharedStrings.xml)');
                        expect(StyleSheet).toHaveBeenCalledWith('PARSED(xl/styles.xml)');

                        expect(Relationships.prototype.findByType).toHaveBeenCalledWith('sharedStrings');
                        expect(ContentTypes.prototype.findByPartName).toHaveBeenCalledWith("/xl/sharedStrings.xml");

                        expect(workbook._zip.remove).toHaveBeenCalledWith("xl/calcChain.xml");
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
    });
});
