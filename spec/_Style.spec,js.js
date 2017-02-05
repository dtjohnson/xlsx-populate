"use strict";

/* eslint camelcase:off */

const proxyquire = require("proxyquire").noCallThru();

describe("_Style", () => {
    let _Style, style, styleSheet, id, xfNode, fontNode, fillNode, borderNode;

    beforeEach(() => {
        _Style = proxyquire("../lib/_Style", {});
        styleSheet = {};
        id = "ID";
        xfNode = {};
        fontNode = {};
        fillNode = {};
        borderNode = {};
        style = new _Style(styleSheet, id, xfNode, fontNode, fillNode, borderNode);
    });

    describe("style", () => {
        it("should get the style with the given name", () => {
            style._get_foo = jasmine.createSpy("_get_foo").and.returnValue("FOO");
            expect(style.style("foo")).toBe("FOO");
            expect(style._get_foo).toHaveBeenCalledWith();
        });

        it("should set the style with the given name", () => {
            style._set_foo = jasmine.createSpy("_set_foo");
            expect(style.style("foo", "FOO")).toBe(style);
            expect(style._set_foo).toHaveBeenCalledWith("FOO");
        });
    });

    describe("bold", () => {
        it("should get/set whether the cell is bold", () => {
            expect(style.style("bold")).toBe(false);
            style.style("bold", true);
            expect(style.style("bold")).toBe(true);
            expect(fontNode).toEqualJson({ b: [{}] });
            style.style("bold", false);
            expect(style.style("bold")).toBe(false);
            expect(fontNode).toEqualJson({});
        });
    });

    describe("italic", () => {
        it("should get/set whether the cell is italic", () => {
            expect(style.style("italic")).toBe(false);
            style.style("italic", true);
            expect(style.style("italic")).toBe(true);
            expect(fontNode).toEqualJson({ i: [{}] });
            style.style("italic", false);
            expect(style.style("italic")).toBe(false);
            expect(fontNode).toEqualJson({});
        });
    });

    describe("underline", () => {
        it("should get/set whether the cell is underline", () => {
            expect(style.style("underline")).toBe(false);
            style.style("underline", true);
            expect(style.style("underline")).toBe(true);
            expect(fontNode).toEqualJson({ u: [{}] });
            style.style("underline", "double");
            expect(style.style("underline")).toBe("double");
            expect(fontNode).toEqualJson({ u: [{ $: { val: "double" } }] });
            style.style("underline", false);
            expect(style.style("underline")).toBe(false);
            expect(fontNode).toEqualJson({});
        });
    });

    describe("strikethrough", () => {
        it("should get/set whether the cell is strikethrough", () => {
            expect(style.style("strikethrough")).toBe(false);
            style.style("strikethrough", true);
            expect(style.style("strikethrough")).toBe(true);
            expect(fontNode).toEqualJson({ strike: [{}] });
            style.style("strikethrough", false);
            expect(style.style("strikethrough")).toBe(false);
            expect(fontNode).toEqualJson({});
        });
    });

    describe("subscript", () => {
        it("should get/set whether the cell is subscript", () => {
            expect(style.style("subscript")).toBe(false);
            style.style("subscript", true);
            expect(style.style("subscript")).toBe(true);
            expect(fontNode).toEqualJson({ vertAlign: [{ $: { val: "subscript" } }] });
            style.style("subscript", false);
            expect(style.style("subscript")).toBe(false);
            expect(fontNode).toEqualJson({});
        });
    });

    describe("superscript", () => {
        it("should get/set whether the cell is superscript", () => {
            expect(style.style("superscript")).toBe(false);
            style.style("superscript", true);
            expect(style.style("superscript")).toBe(true);
            expect(fontNode).toEqualJson({ vertAlign: [{ $: { val: "superscript" } }] });
            style.style("superscript", false);
            expect(style.style("superscript")).toBe(false);
            expect(fontNode).toEqualJson({});
        });
    });
});
