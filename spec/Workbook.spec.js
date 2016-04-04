"use strict";

var path = require("path");
var proxyquire = require("proxyquire").noCallThru();
var xpath = require('xpath');
var DOMParser = require('xmldom').DOMParser;
var parser = new DOMParser();

describe("Workbook", function () {
    var Workbook, fs, Row, JSZip;

    beforeEach(function () {
        fs = jasmine.createSpyObj("fs", ["readFile", "readFileSync", "writeFile", "writeFileSync"]);
        JSZip = {};
        Row = {};
        Workbook = proxyquire("../lib/Workbook", { fs: fs, jszip: JSZip, './Row': Row });
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

    describe("instance", function () {
        beforeEach(function () {

        });

        describe("constructor", function () {
            // TODO
        });

        describe("getSheet", function () {
            // TODO
        });

        describe("getNamedCell", function () {
            // TODO
        });

        describe("output", function () {
            // TODO
        });
    });

    describe("toFile", function () {
        it("should call writeFile with the output", function () {
            Workbook.prototype._initialize = jasmine.createSpy("_initialize");
            Workbook.prototype.output = jasmine.createSpy("output").and.returnValue("some output");
            var cb = function () {};
            var workbook = new Workbook();
            workbook.toFile("some/path.xlsx", cb);
            expect(fs.writeFile).toHaveBeenCalledWith("some/path.xlsx", "some output", cb);
        });
    });

    describe("toFileSync", function () {
        it("should call writeFileSync with the output", function () {
            Workbook.prototype._initialize = jasmine.createSpy("_initialize");
            Workbook.prototype.output = jasmine.createSpy("output").and.returnValue("some output");
            var workbook = new Workbook();
            workbook.toFileSync("some/path.xlsx");
            expect(fs.writeFileSync).toHaveBeenCalledWith("some/path.xlsx", "some output");
        });
    });
});
