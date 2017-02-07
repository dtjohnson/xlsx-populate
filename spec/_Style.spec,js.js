"use strict";

/* eslint camelcase:off */

const proxyquire = require("proxyquire").noCallThru();

describe("_Style", () => {
    let _Style, style, styleSheet, id, xfNode, fontNode, fillNode, borderNode, emptyBorderNode;

    beforeEach(() => {
        _Style = proxyquire("../lib/_Style", {});
        styleSheet = {};
        id = "ID";
        xfNode = {};
        fontNode = {};
        fillNode = {};
        borderNode = { left: [{}], right: [{}], top: [{}], bottom: [{}], diagonal: [{}] };
        emptyBorderNode = { left: [{}], right: [{}], top: [{}], bottom: [{}], diagonal: [{}] };
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
        it("should get/set bold", () => {
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
        it("should get/set italic", () => {
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
        it("should get/set underline", () => {
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
        it("should get/set strikethrough", () => {
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
        it("should get/set subscript", () => {
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
        it("should get/set superscript", () => {
            expect(style.style("superscript")).toBe(false);
            style.style("superscript", true);
            expect(style.style("superscript")).toBe(true);
            expect(fontNode).toEqualJson({ vertAlign: [{ $: { val: "superscript" } }] });
            style.style("superscript", false);
            expect(style.style("superscript")).toBe(false);
            expect(fontNode).toEqualJson({});
        });
    });

    describe("fontSize", () => {
        it("should get/set fontSize", () => {
            expect(style.style("fontSize")).toBe(undefined);
            style.style("fontSize", 17);
            expect(style.style("fontSize")).toBe(17);
            expect(fontNode).toEqualJson({ sz: [{ $: { val: 17 } }] });
            style.style("fontSize", undefined);
            expect(style.style("fontSize")).toBe(undefined);
            expect(fontNode).toEqualJson({});
        });
    });

    describe("fontFamily", () => {
        it("should get/set fontFamily", () => {
            expect(style.style("fontFamily")).toBe(undefined);
            style.style("fontFamily", "Comic Sans MS");
            expect(style.style("fontFamily")).toBe("Comic Sans MS");
            expect(fontNode).toEqualJson({ name: [{ $: { val: "Comic Sans MS" } }] });
            style.style("fontFamily", undefined);
            expect(style.style("fontFamily")).toBe(undefined);
            expect(fontNode).toEqualJson({});
        });
    });

    describe("fontColor", () => {
        it("should get/set fontColor", () => {
            expect(style.style("fontColor")).toBe(undefined);

            style.style("fontColor", "ff0000");
            expect(style.style("fontColor")).toEqualJson({ rgb: "FF0000" });
            expect(fontNode).toEqualJson({ color: [{ $: { rgb: "FF0000" } }] });

            style.style("fontColor", 5);
            expect(style.style("fontColor")).toEqualJson({ theme: 5 });
            expect(fontNode).toEqualJson({ color: [{ $: { theme: 5 } }] });

            style.style("fontColor", { theme: 3, tint: -0.2 });
            expect(style.style("fontColor")).toEqualJson({ theme: 3, tint: -0.2 });
            expect(fontNode).toEqualJson({ color: [{ $: { theme: 3, tint: -0.2 } }] });

            style.style("fontColor", undefined);
            expect(style.style("fontColor")).toBe(undefined);
            expect(fontNode).toEqualJson({});

            fontNode.color = [{ $: { indexed: 7 } }];
            expect(style.style("fontColor")).toEqualJson({ rgb: "00FFFF" });
        });
    });

    describe("horizontalAlignment", () => {
        it("should get/set horizontalAlignment", () => {
            expect(style.style("horizontalAlignment")).toBe(undefined);
            style.style("horizontalAlignment", "center");
            expect(style.style("horizontalAlignment")).toBe("center");
            expect(xfNode).toEqualJson({ alignment: [{ $: { horizontal: "center" } }] });
            style.style("horizontalAlignment", undefined);
            expect(style.style("horizontalAlignment")).toBe(undefined);
            expect(fontNode).toEqualJson({});
        });
    });

    describe("justifyLastLine", () => {
        it("should get/set justifyLastLine", () => {
            expect(style.style("justifyLastLine")).toBe(false);
            style.style("justifyLastLine", true);
            expect(style.style("justifyLastLine")).toBe(true);
            expect(xfNode).toEqualJson({ alignment: [{ $: { justifyLastLine: 1 } }] });
            style.style("justifyLastLine", false);
            expect(style.style("justifyLastLine")).toBe(false);
            expect(fontNode).toEqualJson({});
        });
    });

    describe("indent", () => {
        it("should get/set indent", () => {
            expect(style.style("indent")).toBe(undefined);
            style.style("indent", 3);
            expect(style.style("indent")).toBe(3);
            expect(xfNode).toEqualJson({ alignment: [{ $: { indent: 3 } }] });
            style.style("indent", undefined);
            expect(style.style("indent")).toBe(undefined);
            expect(fontNode).toEqualJson({});
        });
    });

    describe("verticalAlignment", () => {
        it("should get/set verticalAlignment", () => {
            expect(style.style("verticalAlignment")).toBe(undefined);
            style.style("verticalAlignment", "center");
            expect(style.style("verticalAlignment")).toBe("center");
            expect(xfNode).toEqualJson({ alignment: [{ $: { vertical: "center" } }] });
            style.style("verticalAlignment", undefined);
            expect(style.style("verticalAlignment")).toBe(undefined);
            expect(fontNode).toEqualJson({});
        });
    });

    describe("wrapText", () => {
        it("should get/set wrapText", () => {
            expect(style.style("wrapText")).toBe(false);
            style.style("wrapText", true);
            expect(style.style("wrapText")).toBe(true);
            expect(xfNode).toEqualJson({ alignment: [{ $: { wrapText: 1 } }] });
            style.style("wrapText", false);
            expect(style.style("wrapText")).toBe(false);
            expect(fontNode).toEqualJson({});
        });
    });

    describe("shrinkToFit", () => {
        it("should get/set shrinkToFit", () => {
            expect(style.style("shrinkToFit")).toBe(false);
            style.style("shrinkToFit", true);
            expect(style.style("shrinkToFit")).toBe(true);
            expect(xfNode).toEqualJson({ alignment: [{ $: { shrinkToFit: 1 } }] });
            style.style("shrinkToFit", false);
            expect(style.style("shrinkToFit")).toBe(false);
            expect(fontNode).toEqualJson({});
        });
    });

    describe("textDirection", () => {
        it("should get/set textDirection", () => {
            expect(style.style("textDirection")).toBe(undefined);
            style.style("textDirection", "left-to-right");
            expect(style.style("textDirection")).toBe("left-to-right");
            expect(xfNode).toEqualJson({ alignment: [{ $: { readingOrder: 1 } }] });
            style.style("textDirection", "right-to-left");
            expect(style.style("textDirection")).toBe("right-to-left");
            expect(xfNode).toEqualJson({ alignment: [{ $: { readingOrder: 2 } }] });
            style.style("textDirection", undefined);
            expect(style.style("textDirection")).toBe(undefined);
            expect(fontNode).toEqualJson({});
        });
    });

    describe("textRotation", () => {
        it("should get/set indent", () => {
            expect(style.style("textRotation")).toBe(undefined);
            style.style("textRotation", 15);
            expect(style.style("textRotation")).toBe(15);
            expect(xfNode).toEqualJson({ alignment: [{ $: { textRotation: 15 } }] });
            style.style("textRotation", -25);
            expect(style.style("textRotation")).toBe(-25);
            expect(xfNode).toEqualJson({ alignment: [{ $: { textRotation: 115 } }] });
            style.style("textRotation", undefined);
            expect(style.style("textRotation")).toBe(undefined);
            expect(fontNode).toEqualJson({});
        });
    });

    describe("angleTextCounterclockwise", () => {
        it("should get/set angleTextCounterclockwise", () => {
            expect(style.style("angleTextCounterclockwise")).toBe(false);
            style.style("angleTextCounterclockwise", true);
            expect(style.style("angleTextCounterclockwise")).toBe(true);
            expect(xfNode).toEqualJson({ alignment: [{ $: { textRotation: 45 } }] });
            style.style("angleTextCounterclockwise", false);
            expect(style.style("angleTextCounterclockwise")).toBe(false);
            expect(fontNode).toEqualJson({});
        });
    });

    describe("angleTextClockwise", () => {
        it("should get/set angleTextClockwise", () => {
            expect(style.style("angleTextClockwise")).toBe(false);
            style.style("angleTextClockwise", true);
            expect(style.style("angleTextClockwise")).toBe(true);
            expect(xfNode).toEqualJson({ alignment: [{ $: { textRotation: 90 + 45 } }] });
            style.style("angleTextClockwise", false);
            expect(style.style("angleTextClockwise")).toBe(false);
            expect(fontNode).toEqualJson({});
        });
    });

    describe("rotateTextUp", () => {
        it("should get/set rotateTextUp", () => {
            expect(style.style("rotateTextUp")).toBe(false);
            style.style("rotateTextUp", true);
            expect(style.style("rotateTextUp")).toBe(true);
            expect(xfNode).toEqualJson({ alignment: [{ $: { textRotation: 90 } }] });
            style.style("rotateTextUp", false);
            expect(style.style("rotateTextUp")).toBe(false);
            expect(fontNode).toEqualJson({});
        });
    });

    describe("rotateTextDown", () => {
        it("should get/set rotateTextDown", () => {
            expect(style.style("rotateTextDown")).toBe(false);
            style.style("rotateTextDown", true);
            expect(style.style("rotateTextDown")).toBe(true);
            expect(xfNode).toEqualJson({ alignment: [{ $: { textRotation: 90 + 90 } }] });
            style.style("rotateTextDown", false);
            expect(style.style("rotateTextDown")).toBe(false);
            expect(fontNode).toEqualJson({});
        });
    });

    describe("verticalText", () => {
        it("should get/set verticalText", () => {
            expect(style.style("verticalText")).toBe(false);
            style.style("verticalText", true);
            expect(style.style("verticalText")).toBe(true);
            expect(xfNode).toEqualJson({ alignment: [{ $: { textRotation: 255 } }] });
            style.style("verticalText", false);
            expect(style.style("verticalText")).toBe(false);
            expect(fontNode).toEqualJson({});
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
            expect(fillNode).toEqualJson({
                patternFill: [{
                    $: { patternType: "solid" },
                    fgColor: [{
                        $: { rgb: "FF0000" }
                    }]
                }]
            });

            style.style("fill", 5);
            expect(style.style("fill")).toEqualJson({
                type: "solid",
                color: { theme: 5 }
            });
            expect(fillNode).toEqualJson({
                patternFill: [{
                    $: { patternType: "solid" },
                    fgColor: [{
                        $: { theme: 5 }
                    }]
                }]
            });

            style.style("fill", {
                theme: 6,
                tint: -0.25
            });
            expect(style.style("fill")).toEqualJson({
                type: "solid",
                color: { theme: 6, tint: -0.25 }
            });
            expect(fillNode).toEqualJson({
                patternFill: [{
                    $: { patternType: "solid" },
                    fgColor: [{
                        $: { theme: 6, tint: -0.25 }
                    }]
                }]
            });

            style.style("fill", {
                type: "solid",
                color: { rgb: "ff00ff", tint: 0.7 }
            });
            expect(style.style("fill")).toEqualJson({
                type: "solid",
                color: { rgb: "FF00FF", tint: 0.7 }
            });
            expect(fillNode).toEqualJson({
                patternFill: [{
                    $: { patternType: "solid" },
                    fgColor: [{
                        $: { rgb: "FF00FF", tint: 0.7 }
                    }]
                }]
            });

            style.style("fill", undefined);
            expect(style.style("fill")).toBe(undefined);
            expect(fillNode).toEqualJson({});
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
            expect(fillNode).toEqualJson({
                patternFill: [{
                    $: { patternType: "darkVertical" },
                    fgColor: [{
                        $: { rgb: "FF0000" }
                    }],
                    bgColor: [{
                        $: { theme: 7 }
                    }]
                }]
            });


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
            expect(fillNode).toEqualJson({
                patternFill: [{
                    $: { patternType: "gray0625" },
                    fgColor: [{
                        $: { rgb: "AA0000", tint: -1 }
                    }],
                    bgColor: [{
                        $: { theme: 3, tint: 1 }
                    }]
                }]
            });

            style.style("fill", undefined);
            expect(style.style("fill")).toBe(undefined);
            expect(fillNode).toEqualJson({});
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
            expect(fillNode).toEqualJson({
                gradientFill: [{
                    $: { degree: 27 },
                    stop: [
                        {
                            $: { position: 0 },
                            color: [{ $: { rgb: "FFFFFF" } }]
                        },
                        {
                            $: { position: 0.5 },
                            color: [{ $: { theme: 7 } }]
                        },
                        {
                            $: { position: 1 },
                            color: [{ $: { rgb: "000000", tint: 0.5 } }]
                        }
                    ]
                }]
            });

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
            expect(fillNode).toEqualJson({
                gradientFill: [{
                    $: {
                        type: "path",
                        top: 0.1,
                        bottom: 0.2,
                        left: 0.3,
                        right: 0.4
                    },
                    stop: [
                        {
                            $: { position: 0 },
                            color: [{ $: { theme: 0, tint: -0.3 } }]
                        },
                        {
                            $: { position: 1 },
                            color: [{ $: { rgb: "ACACAC" } }]
                        }
                    ]
                }]
            });

            style.style("fill", undefined);
            expect(style.style("fill")).toBe(undefined);
            expect(fillNode).toEqualJson({});
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
                    left: [{ $: { style: "thin" } }],
                    right: [{ $: { style: "thin" } }],
                    top: [{ $: { style: "thin" } }],
                    bottom: [{ $: { style: "thin" } }],
                    diagonal: [{}]
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
                    left: [{ $: { style: "medium" }, color: [{ $: { rgb: "ACACAC" } }] }],
                    right: [{ $: { style: "medium" }, color: [{ $: { rgb: "ACACAC" } }] }],
                    top: [{ $: { style: "medium" }, color: [{ $: { rgb: "ACACAC" } }] }],
                    bottom: [{ $: { style: "medium" }, color: [{ $: { rgb: "ACACAC" } }] }],
                    diagonal: [{}]
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
                    left: [{ color: [{ $: { theme: 1 } }] }],
                    right: [{ color: [{ $: { rgb: "FF0000" } }] }],
                    bottom: [{}],
                    diagonal: [{}],
                    top: [{}]
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
                    left: [{ color: [{ $: { rgb: "FF0000" } }] }],
                    right: [{ color: [{ $: { rgb: "FF0000" } }] }],
                    top: [{ color: [{ $: { rgb: "FF0000" } }] }],
                    bottom: [{ color: [{ $: { rgb: "FF0000" } }] }],
                    diagonal: [{ color: [{ $: { rgb: "FF0000" } }] }]
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
                    left: [{ color: [{ $: { theme: 0 } }] }],
                    right: [{ color: [{ $: { theme: 0 } }] }],
                    top: [{ color: [{ $: { theme: 0 } }] }],
                    bottom: [{ color: [{ $: { theme: 0 } }] }],
                    diagonal: [{ color: [{ $: { theme: 0 } }] }]
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
                    left: [{ $: { style: "thin" } }],
                    right: [{ $: { style: "thick" } }],
                    bottom: [{}],
                    diagonal: [{}],
                    top: [{}]
                });

                style.style("borderStyle", "dashed");
                expect(style.style("borderStyle")).toEqualJson({
                    left: "dashed",
                    right: "dashed",
                    top: "dashed",
                    bottom: "dashed"
                });
                expect(borderNode).toEqualJson({
                    left: [{ $: { style: "dashed" } }],
                    right: [{ $: { style: "dashed" } }],
                    top: [{ $: { style: "dashed" } }],
                    bottom: [{ $: { style: "dashed" } }],
                    diagonal: [{}]
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
                    $: { diagonalUp: 1 },
                    left: [{}],
                    right: [{}],
                    bottom: [{}],
                    diagonal: [{}],
                    top: [{}]
                });

                style.style("diagonalBorderDirection", "down");
                expect(style.style("diagonalBorderDirection")).toBe("down");
                expect(borderNode).toEqualJson({
                    $: { diagonalDown: 1 },
                    left: [{}],
                    right: [{}],
                    bottom: [{}],
                    diagonal: [{}],
                    top: [{}]
                });

                style.style("diagonalBorderDirection", "both");
                expect(style.style("diagonalBorderDirection")).toBe("both");
                expect(borderNode).toEqualJson({
                    $: { diagonalUp: 1, diagonalDown: 1 },
                    left: [{}],
                    right: [{}],
                    bottom: [{}],
                    diagonal: [{}],
                    top: [{}]
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
                    left: [{}],
                    right: [{}],
                    bottom: [{}],
                    diagonal: [{}],
                    top: [{ $: { style: "thin" } }]
                });

                style.style("bottomBorder", { style: "double", color: 6 });
                expect(style.style("bottomBorder")).toEqualJson({ style: "double", color: { theme: 6 } });
                expect(borderNode).toEqualJson({
                    left: [{}],
                    right: [{}],
                    bottom: [{ $: { style: "double" }, color: [{ $: { theme: 6 } }] }],
                    diagonal: [{}],
                    top: [{ $: { style: "thin" } }]
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
                    left: [{}],
                    right: [{ color: [{ $: { rgb: "FF0000" } }] }],
                    bottom: [{}],
                    diagonal: [{}],
                    top: [{}]
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
                    left: [{ $: { style: "thick" } }],
                    right: [{}],
                    bottom: [{}],
                    diagonal: [{}],
                    top: [{}]
                });

                style.style("leftBorderStyle", undefined);
                expect(style.style("leftBorderStyle")).toBe(undefined);
                expect(borderNode).toEqualJson(emptyBorderNode);
            });
        });
    });
});
