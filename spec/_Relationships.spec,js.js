"use strict";

const proxyquire = require("proxyquire").noCallThru();
const DOMParser = require('xmldom').DOMParser;
const parser = new DOMParser();

describe("_Relationships", () => {
    let _Relationship, _Relationships, relationships, relationshipsText;

    beforeEach(() => {
        _Relationship = jasmine.createSpy("_Relationship");
        _Relationship.prototype = jasmine.createSpyObj("_Relationship.prototype", ["type"]);
        _Relationship.prototype.type.and.returnValue("TYPE");

        _Relationships = proxyquire("../lib/_Relationships", {
            "./_Relationship": _Relationship
        });

        relationshipsText = `
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`;
        relationships = new _Relationships(relationshipsText);
    });

    describe("constructor", () => {
        it("should create a content type for each", () => {
            expect(_Relationship.calls.argsFor(0)[0].toString()).toBe(`<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml" xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`);
            expect(_Relationship.calls.argsFor(1)[0].toString()).toBe(`<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml" xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`);
            expect(_Relationship.calls.argsFor(2)[0].toString()).toBe(`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml" xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`);
            expect(relationships._relationships.length).toBe(3);
        });
    });

    describe("add", () => {
        it("should add a new part", () => {
            spyOn(Date, "now").and.returnValue("ID");
            _Relationship.calls.reset();
            relationships.add("TYPE", "TARGET");
            expect(_Relationship.calls.argsFor(0)[0].toString()).toBe(`<Relationship Id="rIdID" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/TYPE" Target="TARGET"/>`);
            expect(relationships._relationships.length).toBe(4);
        });
    });

    describe("findByType", () => {
        it("should return the relationship if matched", () => {
            expect(relationships.findByType("TYPE")).toBe(relationships._relationships[0]);
        });

        it("should return undefined if not matched", () => {
            expect(relationships.findByType("foo")).toBeUndefined();
        });
    });

    describe("toString", () => {
        it("should export to the XML string", () => {
            expect(relationships.toString().trim()).toBe(relationshipsText.trim());
        });
    });
});
