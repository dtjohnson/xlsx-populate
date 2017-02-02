"use strict";

const proxyquire = require("proxyquire").noCallThru();
const DOMParser = require('xmldom').DOMParser;
const parser = new DOMParser();
const xq = require("../lib/xq");

return;
xdescribe("_StyleSheet", () => {
    let _Style, _StyleSheet, styleSheet, styleSheetText;

    beforeEach(() => {
        _Style = jasmine.createSpy("_Style");
        _StyleSheet = proxyquire("../lib/_StyleSheet", {
            "./_Style": _Style
        });
        styleSheetText = `
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main">
    <fonts count="1" x14ac:knownFonts="1"></fonts>
    <fills count="11"></fills>
    <borders count="10"><border foo="bar"/></borders>
    <cellStyleXfs count="1">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    </cellStyleXfs>
    <cellXfs count="19">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyBorder="1" xfId="0"/>
    </cellXfs>
</styleSheet>`;
        styleSheet = new _StyleSheet(styleSheetText);
    });

    describe("constructor", () => {
        fit("should initialize", () => {
            expect(xq.query(styleSheet._xml.documentElement, {
                fonts: { '@count': 0 },
                fills: { '@count': 0 },
                borders: { '@count': 1 },
                cellXfs: { '@count': 1 }
            })).toBeTruthy();
            /*expect(styleSheet._xml.toString()).toBe(`
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main">
    <fonts count="0" x14ac:knownFonts="1"/>
    <fills count="0"/>
    <borders count="1"><border foo="bar"/></borders>
    <cellStyleXfs count="1">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    </cellStyleXfs>
    <cellXfs count="1">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyBorder="1" xfId="0"/>
    </cellXfs>
</styleSheet>`);*/
        });
    });

    describe("createStyle", () => {
        it("should clone an existing style", () => {
            const style = styleSheet.createStyle(0);
            expect(style).toEqual(jasmine.any(_Style));
            expect(styleSheet._xml.toString()).toBe(`
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main"><numFmts/>
    <fonts count="1" x14ac:knownFonts="1"><font/></fonts>
    <fills count="1"><fill/></fills>
    <borders count="2"><border foo="bar"/><border foo="bar"/></borders>
    <cellStyleXfs count="1">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    </cellStyleXfs>
    <cellXfs count="2">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyBorder="1" xfId="0"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" applyBorder="1" xfId="0" applyFont="1" applyFill="1"/></cellXfs>
</styleSheet>`);
        });

        it("should create a new style", () => {
            const style = styleSheet.createStyle(undefined);
            expect(style).toEqual(jasmine.any(_Style));
            expect(styleSheet._xml.toString()).toBe(`
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main"><numFmts/>
    <fonts count="1" x14ac:knownFonts="1"><font/></fonts>
    <fills count="1"><fill/></fills>
    <borders count="2"><border foo="bar"/><border/></borders>
    <cellStyleXfs count="1">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    </cellStyleXfs>
    <cellXfs count="2">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyBorder="1" xfId="0"/>
    <xf fontId="0" fillId="0" borderId="1" applyFont="1" applyFill="1" applyBorder="1"/></cellXfs>
</styleSheet>`);
        });
    });

    describe("getNumberFormatCode", () => {
        it("should return the index if the string already exists", () => {
            expect(styleSheet.getNumberFormatCode(0)).toBe("General");
            expect(styleSheet.getNumberFormatCode(49)).toBe('@');
        });
    });

    describe("getNumberFormatId", () => {
        it("should return an existing code ID", () => {
            expect(styleSheet.getNumberFormatId("General")).toBe(0);
            expect(styleSheet.getNumberFormatId("@")).toBe(49);
        });

        it("should add a custom format node if code doesn't exist", () => {
            expect(styleSheet._numFmtsNode.toString()).toBe('<numFmts/>');
            expect(styleSheet.getNumberFormatId('foo')).toBe(164);
            expect(styleSheet._numFmtsNode.toString()).toBe('<numFmts><numFmt numFmtId="164" formatCode="foo"/></numFmts>');
        });
    });

    describe("toString", () => {
        it("should export to the XML string", () => {
            styleSheet._xml = parser.parseFromString("<foo/>");
            expect(styleSheet.toString()).toBe('<foo/>');
        });
    });
});
