"use strict";

const proxyquire = require("proxyquire");

describe("StyleSheet", () => {
    let Style, StyleSheet, styleSheet, styleSheetNode;

    beforeEach(() => {
        Style = jasmine.createSpy("_Style");
        StyleSheet = proxyquire("../../lib/StyleSheet", {
            "./Style": Style,
            '@noCallThru': true
        });

        styleSheetNode = {
            name: "styleSheet",
            attributes: {
                xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            },
            children: [{
                name: "fonts",
                attributes: {
                    count: 1,
                    'x14ac:knownFonts': 1
                },
                children: []
            }, {
                name: "fills",
                attributes: {
                    count: 11
                },
                children: []
            }, {
                name: "borders",
                attributes: {
                    count: 10
                },
                children: [{
                    name: "border",
                    attributes: {
                        foo: "bar"
                    },
                    children: []
                }]
            }, {
                name: "cellStyleXfs",
                attributes: {
                    count: 1
                },
                children: [{
                    name: "xf",
                    attributes: {
                        numFmtId: 0,
                        fontId: 0,
                        fillId: 0,
                        borderId: 0
                    },
                    children: []
                }]
            }, {
                name: "cellXfs",
                attributes: {
                    count: 19
                },
                children: [{
                    name: "xf",
                    attributes: {
                        numFmtId: 0,
                        fontId: 0,
                        fillId: 0,
                        borderId: 0,
                        applyBorder: 1,
                        xfId: 0
                    },
                    children: []
                }]
            }]
        };
        styleSheet = new StyleSheet(styleSheetNode);
    });

    describe("createStyle", () => {
        it("should clone an existing style", () => {
            const style = styleSheet.createStyle(0);
            expect(style.constructor).toBe(Style);
            expect(styleSheet._node.children).toEqualJson([{
                name: "numFmts",
                attributes: {},
                children: []
            }, {
                name: "fonts",
                attributes: {
                    'x14ac:knownFonts': 1
                },
                children: [
                    { name: "font", attributes: {}, children: [] }
                ]
            }, {
                name: "fills",
                attributes: {},
                children: [
                    { name: "fill", attributes: {}, children: [] }
                ]
            }, {
                name: "borders",
                attributes: {},
                children: [{
                    name: "border",
                    attributes: {
                        foo: "bar"
                    },
                    children: []
                }, {
                    name: "border",
                    attributes: {
                        foo: "bar"
                    },
                    children: [
                        { name: "left", attributes: {}, children: [] },
                        { name: "right", attributes: {}, children: [] },
                        { name: "top", attributes: {}, children: [] },
                        { name: "bottom", attributes: {}, children: [] },
                        { name: "diagonal", attributes: {}, children: [] }
                    ]
                }]
            }, {
                name: "cellStyleXfs",
                attributes: {
                    count: 1
                },
                children: [{
                    name: "xf",
                    attributes: {
                        numFmtId: 0,
                        fontId: 0,
                        fillId: 0,
                        borderId: 0
                    },
                    children: []
                }]
            }, {
                name: "cellXfs",
                attributes: {},
                children: [{
                    name: "xf",
                    attributes: {
                        numFmtId: 0,
                        fontId: 0,
                        fillId: 0,
                        borderId: 0,
                        applyBorder: 1,
                        xfId: 0
                    },
                    children: []
                }, {
                    name: "xf",
                    attributes: {
                        numFmtId: 0,
                        fontId: 0,
                        fillId: 0,
                        borderId: 1,
                        applyFill: 1,
                        applyFont: 1,
                        applyBorder: 1,
                        xfId: 0
                    },
                    children: []
                }]
            }]);
            expect(Style).toHaveBeenCalledWith(
                styleSheet,
                1,
                styleSheet._cellXfsNode.children[1],
                styleSheet._fontsNode.children[0],
                styleSheet._fillsNode.children[0],
                styleSheet._bordersNode.children[1]
            );
        });

        it("should create a new style", () => {
            const style = styleSheet.createStyle(undefined);
            expect(style.constructor).toBe(Style);
            expect(styleSheet._node.children).toEqualJson([{
                name: "numFmts",
                attributes: {},
                children: []
            }, {
                name: "fonts",
                attributes: {
                    'x14ac:knownFonts': 1
                },
                children: [
                    { name: "font", attributes: {}, children: [] }
                ]
            }, {
                name: "fills",
                attributes: {},
                children: [
                    { name: "fill", attributes: {}, children: [] }
                ]
            }, {
                name: "borders",
                attributes: {},
                children: [{
                    name: "border",
                    attributes: {
                        foo: "bar"
                    },
                    children: []
                }, {
                    name: "border",
                    attributes: {},
                    children: [
                        { name: "left", attributes: {}, children: [] },
                        { name: "right", attributes: {}, children: [] },
                        { name: "top", attributes: {}, children: [] },
                        { name: "bottom", attributes: {}, children: [] },
                        { name: "diagonal", attributes: {}, children: [] }
                    ]
                }]
            }, {
                name: "cellStyleXfs",
                attributes: {
                    count: 1
                },
                children: [{
                    name: "xf",
                    attributes: {
                        numFmtId: 0,
                        fontId: 0,
                        fillId: 0,
                        borderId: 0
                    },
                    children: []
                }]
            }, {
                name: "cellXfs",
                attributes: {},
                children: [{
                    name: "xf",
                    attributes: {
                        numFmtId: 0,
                        fontId: 0,
                        fillId: 0,
                        borderId: 0,
                        applyBorder: 1,
                        xfId: 0
                    },
                    children: []
                }, {
                    name: "xf",
                    attributes: {
                        fontId: 0,
                        fillId: 0,
                        borderId: 1,
                        applyFill: 1,
                        applyFont: 1,
                        applyBorder: 1
                    },
                    children: []
                }]
            }]);
            expect(Style).toHaveBeenCalledWith(
                styleSheet,
                1,
                styleSheet._cellXfsNode.children[1],
                styleSheet._fontsNode.children[0],
                styleSheet._fillsNode.children[0],
                styleSheet._bordersNode.children[1]
            );
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
            expect(styleSheet._numFmtsNode.children).toEqualJson([]);
            expect(styleSheet.getNumberFormatId('foo')).toBe(164);
            expect(styleSheet._numFmtsNode.children).toEqualJson([{
                name: 'numFmt',
                attributes: {
                    formatCode: "foo",
                    numFmtId: 164
                }
            }]);
        });
    });

    describe("toXml", () => {
        it("should return the node as is", () => {
            expect(styleSheet.toXml()).toBe(styleSheetNode);
        });
    });

    describe("_cacheNumberFormats", () => {
        it("should cache the number formats", () => {
            styleSheet.getNumberFormatId("foo");
            styleSheet._cacheNumberFormats();
            expect(styleSheet._numberFormatCodesById[0]).toBe("General");
            expect(styleSheet._numberFormatCodesById[49]).toBe("@");
            expect(styleSheet._numberFormatCodesById[164]).toBe("foo");
            expect(styleSheet._numberFormatIdsByCode[`General`]).toBe(0);
            expect(styleSheet._numberFormatIdsByCode[`@`]).toBe(49);
            expect(styleSheet._numberFormatIdsByCode[`foo`]).toBe(164);
        });
    });

    describe("_init", () => {
        it("should add the numFmts node and clear the counts", () => {
            expect(styleSheet._numFmtsNode).toEqualJson({ name: "numFmts", attributes: {}, children: [] });
            expect(styleSheet._fontsNode.attributes.count).toBeUndefined();
            expect(styleSheet._fillsNode.attributes.count).toBeUndefined();
            expect(styleSheet._bordersNode.attributes.count).toBeUndefined();
            expect(styleSheet._cellXfsNode.attributes.count).toBeUndefined();
        });
    });
});
