"use strict";

/* eslint camelcase:off */

const proxyquire = require("proxyquire");
const _ = require("lodash");

describe("Style", () => {
    let Style, style, styleSheet, id, xfNode, fontNode, fillNode, borderNode, emptyBorderNode;

    beforeEach(() => {
        Style = proxyquire("../../lib/Style", {
            '@noCallThru': true
        });
        styleSheet = jasmine.createSpyObj("styleSheet", ['getNumberFormatCode', 'getNumberFormatId']);
        id = "ID";
        xfNode = { name: "xf", attributes: {}, children: [] };
        fontNode = { name: "font", attributes: {}, children: [] };
        fillNode = { name: "fill", attributes: {}, children: [] };
        borderNode = {
            name: "border",
            attributes: {},
            children: [
                { name: "left", attributes: {}, children: [] },
                { name: "right", attributes: {}, children: [] },
                { name: "top", attributes: {}, children: [] },
                { name: "bottom", attributes: {}, children: [] },
                { name: "diagonal", attributes: {}, children: [] }
            ]
        };
        emptyBorderNode = _.cloneDeep(borderNode);
        style = new Style(styleSheet, id, xfNode, fontNode, fillNode, borderNode);
    });

    describe("id", () => {
        it("should return the ID", () => {
            expect(style.id()).toBe("ID");
        });
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
        it("should get/set bold", () => {
            expect(style.style("bold")).toBe(false);
            style.style("bold", true);
            expect(style.style("bold")).toBe(true);
            expect(fontNode.children).toEqualJson([{ name: "b", attributes: {}, children: [] }]);
            style.style("bold", false);
            expect(style.style("bold")).toBe(false);
            expect(fontNode.children).toEqualJson([]);
        });
    });

    describe("italic", () => {
        it("should get/set italic", () => {
            expect(style.style("italic")).toBe(false);
            style.style("italic", true);
            expect(style.style("italic")).toBe(true);
            expect(fontNode.children).toEqualJson([{ name: "i", attributes: {}, children: [] }]);
            style.style("italic", false);
            expect(style.style("italic")).toBe(false);
            expect(fontNode.children).toEqualJson([]);
        });
    });

    describe("underline", () => {
        it("should get/set underline", () => {
            expect(style.style("underline")).toBe(false);
            style.style("underline", true);
            expect(style.style("underline")).toBe(true);
            expect(fontNode.children).toEqualJson([{ name: "u", attributes: {}, children: [] }]);
            style.style("underline", "double");
            expect(style.style("underline")).toBe("double");
            expect(fontNode.children).toEqualJson([{ name: "u", attributes: { val: "double" }, children: [] }]);
            style.style("underline", true);
            expect(style.style("underline")).toBe(true);
            expect(fontNode.children).toEqualJson([{ name: "u", attributes: {}, children: [] }]);
            style.style("underline", false);
            expect(style.style("underline")).toBe(false);
            expect(fontNode.children).toEqualJson([]);
        });
    });

    describe("strikethrough", () => {
        it("should get/set strikethrough", () => {
            expect(style.style("strikethrough")).toBe(false);
            style.style("strikethrough", true);
            expect(style.style("strikethrough")).toBe(true);
            expect(fontNode.children).toEqualJson([{ name: 'strike', attributes: {}, children: [] }]);
            style.style("strikethrough", false);
            expect(style.style("strikethrough")).toBe(false);
            expect(fontNode.children).toEqualJson([]);
        });
    });

    describe("subscript", () => {
        it("should get/set subscript", () => {
            expect(style.style("subscript")).toBe(false);
            style.style("subscript", true);
            expect(style.style("subscript")).toBe(true);
            expect(fontNode.children).toEqualJson([{ name: "vertAlign", attributes: { val: "subscript" }, children: [] }]);
            style.style("subscript", false);
            expect(style.style("subscript")).toBe(false);
            expect(fontNode.children).toEqualJson([]);
        });
    });

    describe("superscript", () => {
        it("should get/set superscript", () => {
            expect(style.style("superscript")).toBe(false);
            style.style("superscript", true);
            expect(style.style("superscript")).toBe(true);
            expect(fontNode.children).toEqualJson([{ name: "vertAlign", attributes: { val: "superscript" }, children: [] }]);
            style.style("superscript", false);
            expect(style.style("superscript")).toBe(false);
            expect(fontNode.children).toEqualJson([]);
        });
    });

    describe("fontSize", () => {
        it("should get/set fontSize", () => {
            expect(style.style("fontSize")).toBe(undefined);
            style.style("fontSize", 17);
            expect(style.style("fontSize")).toBe(17);
            expect(fontNode.children).toEqualJson([{ name: 'sz', attributes: { val: 17 }, children: [] }]);
            style.style("fontSize", undefined);
            expect(style.style("fontSize")).toBe(undefined);
            expect(fontNode.children).toEqualJson([]);
        });
    });

    describe("fontFamily", () => {
        it("should get/set fontFamily", () => {
            expect(style.style("fontFamily")).toBe(undefined);
            style.style("fontFamily", "Comic Sans MS");
            expect(style.style("fontFamily")).toBe("Comic Sans MS");
            expect(fontNode.children).toEqualJson([{ name: 'name', attributes: { val: "Comic Sans MS" }, children: [] }]);
            style.style("fontFamily", undefined);
            expect(style.style("fontFamily")).toBe(undefined);
            expect(fontNode.children).toEqualJson([]);
        });
    });

    describe("fontGenericFamily", () => {
        it("should get/set fontGenericFamily", () => {
            expect(style.style("fontGenericFamily")).toBe(undefined);
            style.style("fontGenericFamily", 1);
            expect(style.style("fontGenericFamily")).toBe(1);
            expect(fontNode.children).toEqualJson([{ name: 'family', attributes: { val: 1 }, children: [] }]);
            style.style("fontGenericFamily", undefined);
            expect(style.style("fontGenericFamily")).toBe(undefined);
            expect(fontNode.children).toEqualJson([]);
        });
    });

    describe("fontScheme", () => {
        it("should get/set fontScheme", () => {
            expect(style.style("fontScheme")).toBe(undefined);
            style.style("fontScheme", 'minor');
            expect(style.style("fontScheme")).toBe('minor');
            expect(fontNode.children).toEqualJson([{ name: 'scheme', attributes: { val: 'minor' }, children: [] }]);
            style.style("fontScheme", undefined);
            expect(style.style("fontScheme")).toBe(undefined);
            expect(fontNode.children).toEqualJson([]);
        });
    });

    describe("fontColor", () => {
        it("should get/set fontColor", () => {
            expect(style.style("fontColor")).toBe(undefined);

            style.style("fontColor", "ff0000");
            expect(style.style("fontColor")).toEqualJson({ rgb: "FF0000" });
            expect(fontNode.children).toEqualJson([{ name: 'color', attributes: { rgb: "FF0000" }, children: [] }]);

            style.style("fontColor", 5);
            expect(style.style("fontColor")).toEqualJson({ theme: 5 });
            expect(fontNode.children).toEqualJson([{ name: 'color', attributes: { theme: 5 }, children: [] }]);

            style.style("fontColor", { theme: 3, tint: -0.2 });
            expect(style.style("fontColor")).toEqualJson({ theme: 3, tint: -0.2 });
            expect(fontNode.children).toEqualJson([{ name: 'color', attributes: { theme: 3, tint: -0.2 }, children: [] }]);

            style.style("fontColor", undefined);
            expect(style.style("fontColor")).toBe(undefined);
            expect(fontNode.children).toEqualJson([]);

            fontNode.children = [{ name: 'color', attributes: { indexed: 7 }, children: [] }];
            expect(style.style("fontColor")).toEqualJson({ rgb: "00FFFF" });
        });
    });

    describe("horizontalAlignment", () => {
        it("should get/set horizontalAlignment", () => {
            expect(style.style("horizontalAlignment")).toBe(undefined);
            style.style("horizontalAlignment", "center");
            expect(style.style("horizontalAlignment")).toBe("center");
            expect(xfNode.children).toEqualJson([{ name: "alignment", attributes: { horizontal: "center" }, children: [] }]);
            style.style("horizontalAlignment", undefined);
            expect(style.style("horizontalAlignment")).toBe(undefined);
            expect(xfNode.children).toEqualJson([]);
        });
    });

    describe("justifyLastLine", () => {
        it("should get/set justifyLastLine", () => {
            expect(style.style("justifyLastLine")).toBe(false);
            style.style("justifyLastLine", true);
            expect(style.style("justifyLastLine")).toBe(true);
            expect(xfNode.children).toEqualJson([{ name: 'alignment', attributes: { justifyLastLine: 1 }, children: [] }]);
            style.style("justifyLastLine", false);
            expect(style.style("justifyLastLine")).toBe(false);
            expect(xfNode.children).toEqualJson([]);
        });
    });

    describe("indent", () => {
        it("should get/set indent", () => {
            expect(style.style("indent")).toBe(undefined);
            style.style("indent", 3);
            expect(style.style("indent")).toBe(3);
            expect(xfNode.children).toEqualJson([{ name: 'alignment', attributes: { indent: 3 }, children: [] }]);
            style.style("indent", undefined);
            expect(style.style("indent")).toBe(undefined);
            expect(xfNode.children).toEqualJson([]);
        });
    });

    describe("verticalAlignment", () => {
        it("should get/set verticalAlignment", () => {
            expect(style.style("verticalAlignment")).toBe(undefined);
            style.style("verticalAlignment", "center");
            expect(style.style("verticalAlignment")).toBe("center");
            expect(xfNode.children).toEqualJson([{ name: 'alignment', attributes: { vertical: "center" }, children: [] }]);
            style.style("verticalAlignment", undefined);
            expect(style.style("verticalAlignment")).toBe(undefined);
            expect(xfNode.children).toEqualJson([]);
        });
    });

    describe("wrapText", () => {
        it("should get/set wrapText", () => {
            expect(style.style("wrapText")).toBe(false);
            style.style("wrapText", true);
            expect(style.style("wrapText")).toBe(true);
            expect(xfNode.children).toEqualJson([{ name: 'alignment', attributes: { wrapText: 1 }, children: [] }]);
            style.style("wrapText", false);
            expect(style.style("wrapText")).toBe(false);
            expect(xfNode.children).toEqualJson([]);
        });
    });

    describe("shrinkToFit", () => {
        it("should get/set shrinkToFit", () => {
            expect(style.style("shrinkToFit")).toBe(false);
            style.style("shrinkToFit", true);
            expect(style.style("shrinkToFit")).toBe(true);
            expect(xfNode.children).toEqualJson([{ name: 'alignment', attributes: { shrinkToFit: 1 }, children: [] }]);
            style.style("shrinkToFit", false);
            expect(style.style("shrinkToFit")).toBe(false);
            expect(xfNode.children).toEqualJson([]);
        });
    });

    describe("textDirection", () => {
        it("should get/set textDirection", () => {
            expect(style.style("textDirection")).toBe(undefined);
            style.style("textDirection", "left-to-right");
            expect(style.style("textDirection")).toBe("left-to-right");
            expect(xfNode.children).toEqualJson([{ name: 'alignment', attributes: { readingOrder: 1 }, children: [] }]);
            style.style("textDirection", "right-to-left");
            expect(style.style("textDirection")).toBe("right-to-left");
            expect(xfNode.children).toEqualJson([{ name: 'alignment', attributes: { readingOrder: 2 }, children: [] }]);
            style.style("textDirection", undefined);
            expect(style.style("textDirection")).toBe(undefined);
            expect(xfNode.children).toEqualJson([]);
        });
    });

    describe("textRotation", () => {
        it("should get/set indent", () => {
            expect(style.style("textRotation")).toBe(undefined);
            style.style("textRotation", 15);
            expect(style.style("textRotation")).toBe(15);
            expect(xfNode.children).toEqualJson([{ name: 'alignment', attributes: { textRotation: 15 }, children: [] }]);
            style.style("textRotation", -25);
            expect(style.style("textRotation")).toBe(-25);
            expect(xfNode.children).toEqualJson([{ name: 'alignment', attributes: { textRotation: 115 }, children: [] }]);
            style.style("textRotation", undefined);
            expect(style.style("textRotation")).toBe(undefined);
            expect(xfNode.children).toEqualJson([]);
        });
    });

    describe("angleTextCounterclockwise", () => {
        it("should get/set angleTextCounterclockwise", () => {
            expect(style.style("angleTextCounterclockwise")).toBe(false);
            style.style("angleTextCounterclockwise", true);
            expect(style.style("angleTextCounterclockwise")).toBe(true);
            expect(xfNode.children).toEqualJson([{ name: 'alignment', attributes: { textRotation: 45 }, children: [] }]);
            style.style("angleTextCounterclockwise", false);
            expect(style.style("angleTextCounterclockwise")).toBe(false);
            expect(xfNode.children).toEqualJson([]);
        });
    });

    describe("angleTextClockwise", () => {
        it("should get/set angleTextClockwise", () => {
            expect(style.style("angleTextClockwise")).toBe(false);
            style.style("angleTextClockwise", true);
            expect(style.style("angleTextClockwise")).toBe(true);
            expect(xfNode.children).toEqualJson([{ name: 'alignment', attributes: { textRotation: 90 + 45 }, children: [] }]);
            style.style("angleTextClockwise", false);
            expect(style.style("angleTextClockwise")).toBe(false);
            expect(xfNode.children).toEqualJson([]);
        });
    });

    describe("rotateTextUp", () => {
        it("should get/set rotateTextUp", () => {
            expect(style.style("rotateTextUp")).toBe(false);
            style.style("rotateTextUp", true);
            expect(style.style("rotateTextUp")).toBe(true);
            expect(xfNode.children).toEqualJson([{ name: 'alignment', attributes: { textRotation: 90 }, children: [] }]);
            style.style("rotateTextUp", false);
            expect(style.style("rotateTextUp")).toBe(false);
            expect(xfNode.children).toEqualJson([]);
        });
    });

    describe("rotateTextDown", () => {
        it("should get/set rotateTextDown", () => {
            expect(style.style("rotateTextDown")).toBe(false);
            style.style("rotateTextDown", true);
            expect(style.style("rotateTextDown")).toBe(true);
            expect(xfNode.children).toEqualJson([{ name: 'alignment', attributes: { textRotation: 90 + 90 }, children: [] }]);
            style.style("rotateTextDown", false);
            expect(style.style("rotateTextDown")).toBe(false);
            expect(xfNode.children).toEqualJson([]);
        });
    });

    describe("verticalText", () => {
        it("should get/set verticalText", () => {
            expect(style.style("verticalText")).toBe(false);
            style.style("verticalText", true);
            expect(style.style("verticalText")).toBe(true);
            expect(xfNode.children).toEqualJson([{ name: 'alignment', attributes: { textRotation: 255 }, children: [] }]);
            style.style("verticalText", false);
            expect(style.style("verticalText")).toBe(false);
            expect(xfNode.children).toEqualJson([]);
        });
    });

    describe("fill", () => {
        it("should get/set solid fill", () => {
            expect(style.style("fill")).toBe(undefined);

            style.style("fill", "ff0000");
            expect(style.style("fill")).toEqualJson({
                type: "solid",
                color: { rgb: "FF0000" }
            });
            expect(fillNode.children).toEqualJson([{
                name: 'patternFill',
                attributes: { patternType: "solid" },
                children: [{
                    name: 'fgColor',
                    attributes: { rgb: "FF0000" },
                    children: []
                }]
            }]);

            style.style("fill", 5);
            expect(style.style("fill")).toEqualJson({
                type: "solid",
                color: { theme: 5 }
            });
            expect(fillNode.children).toEqualJson([{
                name: 'patternFill',
                attributes: { patternType: "solid" },
                children: [{
                    name: 'fgColor',
                    attributes: { theme: 5 },
                    children: []
                }]
            }]);

            style.style("fill", {
                theme: 6,
                tint: -0.25
            });
            expect(style.style("fill")).toEqualJson({
                type: "solid",
                color: { theme: 6, tint: -0.25 }
            });
            expect(fillNode.children).toEqualJson([{
                name: 'patternFill',
                attributes: { patternType: "solid" },
                children: [{
                    name: 'fgColor',
                    attributes: { theme: 6, tint: -0.25 },
                    children: []
                }]
            }]);

            style.style("fill", {
                type: "solid",
                color: { rgb: "ff00ff", tint: 0.7 }
            });
            expect(style.style("fill")).toEqualJson({
                type: "solid",
                color: { rgb: "FF00FF", tint: 0.7 }
            });
            expect(fillNode.children).toEqualJson([{
                name: 'patternFill',
                attributes: { patternType: "solid" },
                children: [{
                    name: 'fgColor',
                    attributes: { rgb: "FF00FF", tint: 0.7 },
                    children: []
                }]
            }]);

            style.style("fill", undefined);
            expect(style.style("fill")).toBe(undefined);
            expect(fillNode.children).toEqualJson([]);
        });

        it("should get/set pattern fill", () => {
            expect(style.style("fill")).toBe(undefined);

            style.style("fill", {
                type: "pattern",
                pattern: "darkVertical",
                foreground: "FF0000",
                background: 7
            });
            expect(style.style("fill")).toEqualJson({
                type: "pattern",
                pattern: "darkVertical",
                foreground: {
                    rgb: "FF0000"
                },
                background: {
                    theme: 7
                }
            });
            expect(fillNode.children).toEqualJson([{
                name: 'patternFill',
                attributes: { patternType: "darkVertical" },
                children: [{
                    name: 'fgColor',
                    attributes: { rgb: "FF0000" },
                    children: []
                }, {
                    name: 'bgColor',
                    attributes: { theme: 7 },
                    children: []
                }]
            }]);

            style.style("fill", {
                type: "pattern",
                pattern: "gray0625",
                foreground: { rgb: "aa0000", tint: -1 },
                background: { theme: 3, tint: 1 }
            });
            expect(style.style("fill")).toEqualJson({
                type: "pattern",
                pattern: "gray0625",
                foreground: {
                    rgb: "AA0000",
                    tint: -1
                },
                background: {
                    theme: 3,
                    tint: 1
                }
            });
            expect(fillNode.children).toEqualJson([{
                name: 'patternFill',
                attributes: { patternType: "gray0625" },
                children: [{
                    name: 'fgColor',
                    attributes: { rgb: "AA0000", tint: -1 },
                    children: []
                }, {
                    name: 'bgColor',
                    attributes: { theme: 3, tint: 1 },
                    children: []
                }]
            }]);

            style.style("fill", undefined);
            expect(style.style("fill")).toBe(undefined);
            expect(fillNode.children).toEqualJson([]);
        });

        it("should get/set gradient fill", () => {
            expect(style.style("fill")).toBe(undefined);

            style.style("fill", {
                type: "gradient",
                angle: 27,
                stops: [
                    { position: 0, color: "ffffff" },
                    { position: 0.5, color: 7 },
                    { position: 1, color: { rgb: "000000", tint: 0.5 } }
                ]
            });
            expect(style.style("fill")).toEqualJson({
                type: "gradient",
                gradientType: "linear",
                angle: 27,
                stops: [
                    { position: 0, color: { rgb: "FFFFFF" } },
                    { position: 0.5, color: { theme: 7 } },
                    { position: 1, color: { rgb: "000000", tint: 0.5 } }
                ]
            });
            expect(fillNode.children).toEqualJson([{
                name: 'gradientFill',
                attributes: { degree: 27 },
                children: [
                    {
                        name: 'stop',
                        attributes: { position: 0 },
                        children: [{ name: 'color', attributes: { rgb: "FFFFFF" }, children: [] }]
                    },
                    {
                        name: 'stop',
                        attributes: { position: 0.5 },
                        children: [{ name: 'color', attributes: { theme: 7 }, children: [] }]
                    },
                    {
                        name: 'stop',
                        attributes: { position: 1 },
                        children: [{ name: 'color', attributes: { rgb: "000000", tint: 0.5 }, children: [] }]
                    }
                ]
            }]);

            style.style("fill", {
                type: "gradient",
                gradientType: "path",
                top: 0.1,
                bottom: 0.2,
                left: 0.3,
                right: 0.4,
                stops: [
                    { position: 0, color: { theme: 0, tint: -0.3 } },
                    { position: 1, color: "acacac" }
                ]
            });
            expect(style.style("fill")).toEqualJson({
                type: "gradient",
                gradientType: "path",
                top: 0.1,
                bottom: 0.2,
                left: 0.3,
                right: 0.4,
                stops: [
                    { position: 0, color: { theme: 0, tint: -0.3 } },
                    { position: 1, color: { rgb: "ACACAC" } }
                ]
            });
            expect(fillNode.children).toEqualJson([{
                name: 'gradientFill',
                attributes: {
                    type: "path",
                    top: 0.1,
                    bottom: 0.2,
                    left: 0.3,
                    right: 0.4
                },
                children: [
                    {
                        name: 'stop',
                        attributes: { position: 0 },
                        children: [{ name: 'color', attributes: { theme: 0, tint: -0.3 }, children: [] }]
                    },
                    {
                        name: 'stop',
                        attributes: { position: 1 },
                        children: [{ name: 'color', attributes: { rgb: "ACACAC" }, children: [] }]
                    }
                ]
            }]);

            style.style("fill", undefined);
            expect(style.style("fill")).toBe(undefined);
            expect(fillNode.children).toEqualJson([]);
        });
    });

    describe("border", () => {
        describe("border", () => {
            it("should get/set border", () => {
                expect(style.style("borderColor")).toEqualJson({});
                expect(borderNode).toEqualJson(emptyBorderNode);

                style.style("border", "thin");
                expect(style.style("border")).toEqualJson({
                    left: { style: "thin" },
                    right: { style: "thin" },
                    top: { style: "thin" },
                    bottom: { style: "thin" }
                });
                expect(borderNode).toEqualJson({
                    name: 'border',
                    attributes: {},
                    children: [
                        { name: 'left', attributes: { style: "thin" }, children: [] },
                        { name: 'right', attributes: { style: "thin" }, children: [] },
                        { name: 'top', attributes: { style: "thin" }, children: [] },
                        { name: 'bottom', attributes: { style: "thin" }, children: [] },
                        { name: 'diagonal', attributes: {}, children: [] }
                    ]
                });

                style.style("border", undefined);
                expect(style.style("border")).toEqualJson({});
                expect(borderNode).toEqualJson(emptyBorderNode);

                style.style("border", { style: "medium", color: { rgb: "acacac" } });
                expect(style.style("border")).toEqualJson({
                    left: { style: "medium", color: { rgb: "ACACAC" } },
                    right: { style: "medium", color: { rgb: "ACACAC" } },
                    top: { style: "medium", color: { rgb: "ACACAC" } },
                    bottom: { style: "medium", color: { rgb: "ACACAC" } }
                });
                expect(borderNode).toEqualJson({
                    name: 'border',
                    attributes: {},
                    children: [
                        { name: 'left', attributes: { style: "medium" }, children: [{ name: 'color', attributes: { rgb: "ACACAC" }, children: [] }] },
                        { name: 'right', attributes: { style: "medium" }, children: [{ name: 'color', attributes: { rgb: "ACACAC" }, children: [] }] },
                        { name: 'top', attributes: { style: "medium" }, children: [{ name: 'color', attributes: { rgb: "ACACAC" }, children: [] }] },
                        { name: 'bottom', attributes: { style: "medium" }, children: [{ name: 'color', attributes: { rgb: "ACACAC" }, children: [] }] },
                        { name: 'diagonal', attributes: {}, children: [] }
                    ]
                });

                style.style("border", undefined);
                expect(style.style("border")).toEqualJson({});
                expect(borderNode).toEqualJson(emptyBorderNode);

                style.style("border", {
                    left: { color: 0 },
                    top: "dashed"
                });
                expect(style.style("border")).toEqualJson({
                    left: { color: { theme: 0 } },
                    top: { style: "dashed" }
                });
            });
        });

        describe("borderColor", () => {
            it("should get/set borderColor", () => {
                expect(style.style("borderColor")).toEqualJson({});
                expect(borderNode).toEqualJson(emptyBorderNode);

                style.style("borderColor", {
                    left: 1,
                    right: "ff0000"
                });
                expect(style.style("borderColor")).toEqualJson({
                    left: { theme: 1 },
                    right: { rgb: "FF0000" }
                });
                expect(borderNode).toEqualJson({
                    name: 'border',
                    attributes: {},
                    children: [
                        { name: 'left', attributes: {}, children: [{ name: 'color', attributes: { theme: 1 }, children: [] }] },
                        { name: 'right', attributes: {}, children: [{ name: 'color', attributes: { rgb: "FF0000" }, children: [] }] },
                        { name: 'top', attributes: {}, children: [] },
                        { name: 'bottom', attributes: {}, children: [] },
                        { name: 'diagonal', attributes: {}, children: [] }
                    ]
                });

                style.style("borderColor", "ff0000");
                expect(style.style("borderColor")).toEqualJson({
                    left: { rgb: "FF0000" },
                    right: { rgb: "FF0000" },
                    top: { rgb: "FF0000" },
                    bottom: { rgb: "FF0000" },
                    diagonal: { rgb: "FF0000" }
                });
                expect(borderNode).toEqualJson({
                    name: 'border',
                    attributes: {},
                    children: [
                        { name: 'left', attributes: {}, children: [{ name: 'color', attributes: { rgb: "FF0000" }, children: [] }] },
                        { name: 'right', attributes: {}, children: [{ name: 'color', attributes: { rgb: "FF0000" }, children: [] }] },
                        { name: 'top', attributes: {}, children: [{ name: 'color', attributes: { rgb: "FF0000" }, children: [] }] },
                        { name: 'bottom', attributes: {}, children: [{ name: 'color', attributes: { rgb: "FF0000" }, children: [] }] },
                        { name: 'diagonal', attributes: {}, children: [{ name: 'color', attributes: { rgb: "FF0000" }, children: [] }] }
                    ]
                });

                style.style("borderColor", 0);
                expect(style.style("borderColor")).toEqualJson({
                    left: { theme: 0 },
                    right: { theme: 0 },
                    top: { theme: 0 },
                    bottom: { theme: 0 },
                    diagonal: { theme: 0 }
                });
                expect(borderNode).toEqualJson({
                    name: 'border',
                    attributes: {},
                    children: [
                        { name: 'left', attributes: {}, children: [{ name: 'color', attributes: { theme: 0 }, children: [] }] },
                        { name: 'right', attributes: {}, children: [{ name: 'color', attributes: { theme: 0 }, children: [] }] },
                        { name: 'top', attributes: {}, children: [{ name: 'color', attributes: { theme: 0 }, children: [] }] },
                        { name: 'bottom', attributes: {}, children: [{ name: 'color', attributes: { theme: 0 }, children: [] }] },
                        { name: 'diagonal', attributes: {}, children: [{ name: 'color', attributes: { theme: 0 }, children: [] }] }
                    ]
                });

                style.style("borderColor", undefined);
                expect(style.style("borderColor")).toEqualJson({});
                expect(borderNode).toEqualJson(emptyBorderNode);
            });
        });

        describe("borderStyle", () => {
            it("should get/set borderStyle", () => {
                expect(style.style("borderStyle")).toEqualJson({});
                expect(borderNode).toEqualJson(emptyBorderNode);

                style.style("borderStyle", {
                    left: "thin",
                    right: "thick"
                });
                expect(style.style("borderStyle")).toEqualJson({
                    left: "thin",
                    right: "thick"
                });
                expect(borderNode).toEqualJson({
                    name: 'border',
                    attributes: {},
                    children: [
                        { name: 'left', attributes: { style: "thin" }, children: [] },
                        { name: 'right', attributes: { style: "thick" }, children: [] },
                        { name: 'top', attributes: {}, children: [] },
                        { name: 'bottom', attributes: {}, children: [] },
                        { name: 'diagonal', attributes: {}, children: [] }
                    ]
                });

                style.style("borderStyle", "dashed");
                expect(style.style("borderStyle")).toEqualJson({
                    left: "dashed",
                    right: "dashed",
                    top: "dashed",
                    bottom: "dashed"
                });
                expect(borderNode).toEqualJson({
                    name: 'border',
                    attributes: {},
                    children: [
                        { name: 'left', attributes: { style: "dashed" }, children: [] },
                        { name: 'right', attributes: { style: "dashed" }, children: [] },
                        { name: 'top', attributes: { style: "dashed" }, children: [] },
                        { name: 'bottom', attributes: { style: "dashed" }, children: [] },
                        { name: 'diagonal', attributes: {}, children: [] }
                    ]
                });

                style.style("borderStyle", undefined);
                expect(style.style("borderStyle")).toEqualJson({});
                expect(borderNode).toEqualJson(emptyBorderNode);
            });
        });

        describe("diagonalBorderDirection", () => {
            it("should get/set diagonalBorderDirection", () => {
                expect(style.style("diagonalBorderDirection")).toBe(undefined);
                expect(borderNode).toEqualJson(emptyBorderNode);

                style.style("diagonalBorderDirection", "up");
                expect(style.style("diagonalBorderDirection")).toBe("up");
                expect(borderNode).toEqualJson({
                    name: 'border',
                    attributes: { diagonalUp: 1 },
                    children: [
                        { name: 'left', attributes: {}, children: [] },
                        { name: 'right', attributes: {}, children: [] },
                        { name: 'top', attributes: {}, children: [] },
                        { name: 'bottom', attributes: {}, children: [] },
                        { name: 'diagonal', attributes: {}, children: [] }
                    ]
                });

                style.style("diagonalBorderDirection", "down");
                expect(style.style("diagonalBorderDirection")).toBe("down");
                expect(borderNode).toEqualJson({
                    name: 'border',
                    attributes: { diagonalDown: 1 },
                    children: [
                        { name: 'left', attributes: {}, children: [] },
                        { name: 'right', attributes: {}, children: [] },
                        { name: 'top', attributes: {}, children: [] },
                        { name: 'bottom', attributes: {}, children: [] },
                        { name: 'diagonal', attributes: {}, children: [] }
                    ]
                });

                style.style("diagonalBorderDirection", "both");
                expect(style.style("diagonalBorderDirection")).toBe("both");
                expect(borderNode).toEqualJson({
                    name: 'border',
                    attributes: { diagonalUp: 1, diagonalDown: 1 },
                    children: [
                        { name: 'left', attributes: {}, children: [] },
                        { name: 'right', attributes: {}, children: [] },
                        { name: 'top', attributes: {}, children: [] },
                        { name: 'bottom', attributes: {}, children: [] },
                        { name: 'diagonal', attributes: {}, children: [] }
                    ]
                });

                style.style("diagonalBorderDirection", undefined);
                expect(style.style("diagonalBorderDirection")).toBe(undefined);
                expect(borderNode).toEqualJson(emptyBorderNode);
            });
        });

        describe("sideBorder", () => {
            it("should get/set sideBorder", () => {
                expect(style.style("topBorder")).toBe(undefined);
                expect(borderNode).toEqualJson(emptyBorderNode);

                style.style("topBorder", "thin");
                expect(style.style("topBorder")).toEqualJson({ style: "thin" });
                expect(borderNode).toEqualJson({
                    name: 'border',
                    attributes: {},
                    children: [
                        { name: 'left', attributes: {}, children: [] },
                        { name: 'right', attributes: {}, children: [] },
                        { name: 'top', attributes: { style: "thin" }, children: [] },
                        { name: 'bottom', attributes: {}, children: [] },
                        { name: 'diagonal', attributes: {}, children: [] }
                    ]
                });

                style.style("bottomBorder", { style: "double", color: 6 });
                expect(style.style("bottomBorder")).toEqualJson({ style: "double", color: { theme: 6 } });
                expect(borderNode).toEqualJson({
                    name: 'border',
                    attributes: {},
                    children: [
                        { name: 'left', attributes: {}, children: [] },
                        { name: 'right', attributes: {}, children: [] },
                        { name: 'top', attributes: { style: "thin" }, children: [] },
                        { name: 'bottom', attributes: { style: "double" }, children: [{ name: 'color', attributes: { theme: 6 }, children: [] }] },
                        { name: 'diagonal', attributes: {}, children: [] }
                    ]
                });

                style.style("topBorder", undefined).style("bottomBorder", undefined);
                expect(style.style("topBorder")).toBe(undefined);
                expect(borderNode).toEqualJson(emptyBorderNode);
            });
        });

        describe("sideBorderColor", () => {
            it("should get/set sideBorderColor", () => {
                expect(style.style("rightBorderColor")).toBe(undefined);
                expect(borderNode).toEqualJson(emptyBorderNode);

                style.style("rightBorderColor", "ff0000");
                expect(style.style("rightBorderColor")).toEqualJson({ rgb: "FF0000" });
                expect(borderNode).toEqualJson({
                    name: 'border',
                    attributes: {},
                    children: [
                        { name: 'left', attributes: {}, children: [] },
                        { name: 'right', attributes: {}, children: [{ name: 'color', attributes: { rgb: "FF0000" }, children: [] }] },
                        { name: 'top', attributes: {}, children: [] },
                        { name: 'bottom', attributes: {}, children: [] },
                        { name: 'diagonal', attributes: {}, children: [] }
                    ]
                });

                style.style("rightBorderColor", undefined);
                expect(style.style("rightBorderColor")).toBe(undefined);
                expect(borderNode).toEqualJson(emptyBorderNode);
            });
        });

        describe("sideBorderStyle", () => {
            it("should get/set sideBorderStyle", () => {
                expect(style.style("leftBorderStyle")).toBe(undefined);
                expect(borderNode).toEqualJson(emptyBorderNode);

                style.style("leftBorderStyle", "thick");
                expect(style.style("leftBorderStyle")).toBe("thick");
                expect(borderNode).toEqualJson({
                    name: 'border',
                    attributes: {},
                    children: [
                        { name: 'left', attributes: { style: "thick" }, children: [] },
                        { name: 'right', attributes: {}, children: [] },
                        { name: 'top', attributes: {}, children: [] },
                        { name: 'bottom', attributes: {}, children: [] },
                        { name: 'diagonal', attributes: {}, children: [] }
                    ]
                });

                style.style("leftBorderStyle", undefined);
                expect(style.style("leftBorderStyle")).toBe(undefined);
                expect(borderNode).toEqualJson(emptyBorderNode);
            });
        });
    });

    describe("numberFormat", () => {
        it("should get/set numberFormat", () => {
            styleSheet.getNumberFormatCode.and.returnValue("foo");
            styleSheet.getNumberFormatId.and.returnValue(7);

            expect(style.style("numberFormat")).toBe("foo");
            expect(styleSheet.getNumberFormatCode).toHaveBeenCalledWith(0);

            style.style("numberFormat", "bar");
            expect(styleSheet.getNumberFormatId).toHaveBeenCalledWith('bar');
            expect(xfNode).toEqualJson({ name: "xf", attributes: { numFmtId: 7 }, children: [] });
            expect(style.style("numberFormat")).toBe("foo");
            expect(styleSheet.getNumberFormatCode).toHaveBeenCalledWith(7);
        });
    });
});
