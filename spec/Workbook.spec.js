"use strict";

var path = require("path");
var proxyquire = require("proxyquire").noCallThru();
var xpath = require('xpath');
var DOMParser = require('xmldom').DOMParser;
var parser = new DOMParser();

describe("Workbook", function () {
    var Workbook, fs, JSZip, Sheet;

    beforeEach(function () {
        fs = jasmine.createSpyObj("fs", ["readFile", "readFileSync", "writeFile", "writeFileSync"]);
        JSZip = jasmine.createSpy("JSZip");
        Sheet = jasmine.createSpy("Sheet");
        Workbook = proxyquire("../lib/Workbook", { fs: fs, jszip: JSZip, './Sheet': Sheet });
    });

    describe("static", function () {
        beforeEach(function () {
            Workbook.prototype._initialize = jasmine.createSpy("_initialize");
        });

        describe("fromFileSync", function () {
            it("should create a workbook from the file data", function () {
                var data = {};
                fs.readFileSync.and.returnValue(data);
                var workbook = Workbook.fromFileSync("some/path.xlsx");
                expect(fs.readFileSync).toHaveBeenCalledWith("some/path.xlsx");
                expect(Workbook.prototype._initialize).toHaveBeenCalledWith(data);
                expect(workbook instanceof Workbook).toBe(true);
            });
        });

        describe("fromFile", function () {
            it("should create a workbook from the file data", function () {
                var data = {};
                var cb = jasmine.createSpy("cb").and.callFake(function (err, workbook) {
                    expect(fs.readFile).toHaveBeenCalledWith("some/path.xlsx", jasmine.any(Function));
                    expect(err).toBeFalsy();
                    expect(Workbook.prototype._initialize).toHaveBeenCalledWith(data);
                    expect(workbook instanceof Workbook).toBe(true);
                });
                fs.readFile.and.callFake(function (path, cb) {
                    cb(null, data);
                });

                Workbook.fromFile("some/path.xlsx", cb);
                expect(cb).toHaveBeenCalled();
            });
        });

        describe("fromBlankSync", function () {
            it("should call fromFileSync with the blank workbook path", function () {
                Workbook.fromFileSync = jasmine.createSpy("fromFileSync");
                var workbook = Workbook.fromBlankSync();
                expect(Workbook.fromFileSync).toHaveBeenCalledWith(path.join(__dirname, "../lib/blank.xlsx"));
            });
        });

        describe("fromBlank", function () {
            it("should call fromFile with the blank workbook path", function () {
                Workbook.fromFile = jasmine.createSpy("fromFile");
                var cb = function () {};
                Workbook.fromBlank(cb);
                expect(Workbook.fromFile).toHaveBeenCalledWith(path.join(__dirname, "../lib/blank.xlsx"), cb);
            });
        });
    });

    describe("constructor", function () {
        it("should initialize the workbook and create the sheet objects", function () {
            var workbookText = '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets><sheet name="Tom"/><sheet name="Jerry"/></sheets></workbook>';
            var relsText = '<Relationships/>';
            var sheetText = [
                '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>56</v></c></row></sheetData></worksheet>',
                '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>56</v><f>7*8</f></c></row></sheetData></worksheet>'
            ];

            var files = {
                "xl/workbook.xml": workbookText,
                "xl/_rels/workbook.xml.rels": relsText,
                "xl/worksheets/sheet1.xml": sheetText[0],
                "xl/worksheets/sheet2.xml": sheetText[1]
            };

            JSZip.prototype.file = function (fileName) {
                return {
                    asText: function () {
                        return files[fileName];
                    }
                };
            };

            var data = {};
            var workbook = new Workbook(data);

            expect(workbook._workbookXML.toString()).toBe(workbookText);
            expect(workbook._relsXML.toString()).toBe(relsText);
            expect(workbook._sheetsNode.toString()).toBe('<sheets><sheet name="Tom"/><sheet name="Jerry"/></sheets>');
            expect(workbook._sheets.length).toBe(2);
            var firstCallArgs = Sheet.calls.first().args;
            expect(firstCallArgs[0]).toBe(workbook);
            expect(firstCallArgs[1].toString()).toBe('<sheet name="Tom"/>');
            expect(firstCallArgs[2].toString()).toBe(sheetText[0]);

            var secondCallArgs = Sheet.calls.mostRecent().args;
            expect(secondCallArgs[0]).toBe(workbook);
            expect(secondCallArgs[1].toString()).toBe('<sheet name="Jerry"/>');
            expect(secondCallArgs[2].toString()).toBe('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f>7*8</f></c></row></sheetData></worksheet>'); // Formula values removed.
        });
    });

    describe("instance", function () {
        beforeEach(function () {
            Workbook.prototype._initialize = jasmine.createSpy("_initialize");
        });

        describe("createSheet", function () {
            var workbook, initialSheet;
            beforeEach(function () {
                initialSheet = {};
                workbook = new Workbook();
                workbook._sheets = [initialSheet];
                workbook._sheetsNode = parser.parseFromString('<sheets xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>').documentElement;
                workbook._relsXML = parser.parseFromString('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>').documentElement;
            });

            it("should throw an error if the sheet index is less than 0", function () {
                expect(function () {
                    workbook.createSheet("foo", -1);
                }).toThrow();
            });

            it("should throw an error if the sheet index is not a number", function () {
                expect(function () {
                    workbook.createSheet("foo", "bar");
                }).toThrow();
            });

            it("should throw an error if the sheet index is greater than the number of sheets", function () {
                expect(function () {
                    workbook.createSheet("foo", 3);
                }).toThrow();
            });

            it("should create a new sheet at the end of the workbook", function () {
                var sheet = workbook.createSheet("foo");
                expect(workbook._sheetsNode.toString()).toBe('<sheets xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheet name="Sheet1" sheetId="1" r:id="xpopId1"/><sheet name="foo" sheetId="2" r:id="xpopId2"/></sheets>');
                expect(workbook._sheets).toEqual([initialSheet, sheet]);
                expect(workbook._relsXML.toString()).toBe('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="xpopId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="xpopId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>');

                var args = Sheet.calls.argsFor(0);
                expect(args.length).toBe(3);
                expect(args[0]).toBe(workbook);
                expect(args[1].toString()).toBe('<sheet name="foo" sheetId="2" r:id="xpopId2"/>');
                expect(args[2].toString()).toBe('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>');
            });

            it("should create a new sheet at the beginning of the workbook", function () {
                var sheet = workbook.createSheet("bar", 0);
                expect(workbook._sheetsNode.toString()).toBe('<sheets xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheet name="bar" sheetId="1" r:id="xpopId1"/><sheet name="Sheet1" sheetId="2" r:id="xpopId2"/></sheets>');
                expect(workbook._sheets).toEqual([sheet, initialSheet]);
                expect(workbook._relsXML.toString()).toBe('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="xpopId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="xpopId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>');

                var args = Sheet.calls.argsFor(0);
                expect(args.length).toBe(3);
                expect(args[0]).toBe(workbook);
                expect(args[1].toString()).toBe('<sheet name="bar" sheetId="1" r:id="xpopId1"/>');
                expect(args[2].toString()).toBe('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>');
            });
        });

        describe("getSheet", function () {
            var workbook;
            beforeEach(function () {
                workbook = new Workbook();
                workbook._sheets = [{
                    getName: function () {
                        return "Foo";
                    }
                }, {
                    getName: function () {
                        return "Bar";
                    }
                }];
            });

            it("should return the sheet with the given index", function () {
                expect(workbook.getSheet(0)).toBe(workbook._sheets[0]);
                expect(workbook.getSheet(1)).toBe(workbook._sheets[1]);
            });

            it("should return the sheet with the given name", function () {
                expect(workbook.getSheet("Bar")).toBe(workbook._sheets[1]);
                expect(workbook.getSheet("Foo")).toBe(workbook._sheets[0]);
            });
        });

        describe("getNamedCell", function () {
            var workbook;
            beforeEach(function () {
                workbook = new Workbook();
                workbook._workbookXML = parser.parseFromString('<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><definedNames><definedName name="foo">Sheet2!B3</definedName></definedNames></workbook>').documentElement;
            });

            it("should return the cell with the given name", function () {
                var cell = {};
                var getCell = jasmine.createSpy("getCell").and.returnValue(cell);
                workbook.getSheet = jasmine.createSpy("getSheet").and.returnValue({ getCell: getCell });

                var c = workbook.getNamedCell("foo");
                expect(c).toBe(cell);
                expect(getCell).toHaveBeenCalledWith(3, 2);
                expect(workbook.getSheet).toHaveBeenCalledWith("Sheet2");
            });

            it("should return undefined if no matching cell found", function () {
                var c = workbook.getNamedCell("bar");
                expect(c).toBeUndefined();
            });
        });

        describe("output", function () {
            it("should output the XML", function () {
                var workbook = new Workbook();
                workbook._workbookXML = parser.parseFromString('<workbook/>').documentElement;
                workbook._relsXML = parser.parseFromString('<Relationships/>').documentElement;
                workbook._sheets = [{
                    _sheetXML: parser.parseFromString('<sheet id="1"/>').documentElement
                }, {
                    _sheetXML: parser.parseFromString('<sheet id="2"/>').documentElement
                }];

                var generated = {};
                workbook._zip = jasmine.createSpyObj("_zip", ["file", "generate", "remove"]);
                workbook._zip.generate.and.returnValue(generated);

                var output = workbook.output();
                expect(output).toBe(generated);
                expect(workbook._zip.file).toHaveBeenCalledWith("xl/workbook.xml", '<workbook/>');
                expect(workbook._zip.file).toHaveBeenCalledWith("xl/_rels/workbook.xml.rels", '<Relationships/>');
                expect(workbook._zip.file).toHaveBeenCalledWith("xl/worksheets/sheet1.xml", '<sheet id="1"/>');
                expect(workbook._zip.file).toHaveBeenCalledWith("xl/worksheets/sheet2.xml", '<sheet id="2"/>');
                expect(workbook._zip.remove).toHaveBeenCalledWith("xl/calcChain.xml");
                expect(workbook._zip.generate).toHaveBeenCalledWith({ type: "nodebuffer" });
            });
        });

        describe("toFile", function () {
            it("should call writeFile with the output", function () {
                Workbook.prototype.output = jasmine.createSpy("output").and.returnValue("some output");
                var cb = function () {};
                var workbook = new Workbook();
                workbook.toFile("some/path.xlsx", cb);
                expect(fs.writeFile).toHaveBeenCalledWith("some/path.xlsx", "some output", cb);
            });
        });

        describe("toFileSync", function () {
            it("should call writeFileSync with the output", function () {
                Workbook.prototype.output = jasmine.createSpy("output").and.returnValue("some output");
                var workbook = new Workbook();
                workbook.toFileSync("some/path.xlsx");
                expect(fs.writeFileSync).toHaveBeenCalledWith("some/path.xlsx", "some output");
            });
        });
    });
});
