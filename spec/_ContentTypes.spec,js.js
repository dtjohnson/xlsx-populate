"use strict";

const proxyquire = require("proxyquire").noCallThru();
const DOMParser = require('xmldom').DOMParser;
const parser = new DOMParser();

describe("_ContentTypes", () => {
    let _ContentType, _ContentTypes, contentTypes, contentTypesText;

    beforeEach(() => {
        _ContentType = jasmine.createSpy("_ContentType");
        _ContentType.prototype = jasmine.createSpyObj("_ContentType.prototype", ["partName"]);
        _ContentType.prototype.partName.and.returnValue("PART_NAME");

        _ContentTypes = proxyquire("../lib/_ContentTypes", {
            "./_ContentType": _ContentType
        });

        contentTypesText = `
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`;
        contentTypes = new _ContentTypes(contentTypesText);
    });

    describe("constructor", () => {
        it("should create a content type for each", () => {
            expect(_ContentType.calls.argsFor(0)[0].toString()).toBe(`<Default Extension="xml" ContentType="application/xml" xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`);
            expect(_ContentType.calls.argsFor(1)[0].toString()).toBe(`<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`);
            expect(_ContentType.calls.argsFor(2)[0].toString()).toBe(`<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`);
            expect(contentTypes._contentTypes.length).toBe(3);
        });
    });

    describe("add", () => {
        it("should add a new part", () => {
            _ContentType.calls.reset();
            contentTypes.add("NEW_PART_NAME", "NEW_CONTENT_TYPE");
            expect(_ContentType.calls.argsFor(0)[0].toString()).toBe(`<Override PartName="NEW_PART_NAME" ContentType="NEW_CONTENT_TYPE"/>`);
            expect(contentTypes._contentTypes.length).toBe(4);
        });
    });

    describe("findByPartName", () => {
        it("should return the part if matched", () => {
            expect(contentTypes.findByPartName("PART_NAME")).toBe(contentTypes._contentTypes[0]);
        });

        it("should return undefined if not matched", () => {
            expect(contentTypes.findByPartName("foo")).toBeUndefined();
        });
    });

    describe("toString", () => {
        it("should export to the XML string", () => {
            expect(contentTypes.toString().trim()).toBe(contentTypesText.trim());
        });
    });
});
