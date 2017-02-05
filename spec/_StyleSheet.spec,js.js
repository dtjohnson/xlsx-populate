"use strict";

const proxyquire = require("proxyquire").noCallThru();

describe("_StyleSheet", () => {
    let _Style, _StyleSheet, styleSheet, styleSheetNode;

    beforeEach(() => {
        _Style = jasmine.createSpy("_Style");
        _StyleSheet = proxyquire("../lib/_StyleSheet", {
            "./_Style": _Style
        });

        styleSheetNode = {
            styleSheet: {
                $: {
                    xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                },
                fonts: [{
                    $: {
                        count: 1,
                        'x14ac:knownFonts': 1
                    },
                    font: []
                }],
                fills: [{
                    $: {
                        count: 11
                    },
                    fill: []
                }],
                borders: [{
                    $: {
                        count: 10
                    },
                    border: [{
                        $: {
                            foo: "bar"
                        }
                    }]
                }],
                cellStyleXfs: [{
                    $: {
                        count: 1
                    },
                    xf: [{
                        $: {
                            numFmtId: 0,
                            fontId: 0,
                            fillId: 0,
                            borderId: 0
                        }
                    }]
                }],
                cellXfs: [{
                    $: {
                        count: 19
                    },
                    xf: [{
                        $: {
                            numFmtId: 0,
                            fontId: 0,
                            fillId: 0,
                            borderId: 0,
                            applyBorder: 1,
                            xfId: 0
                        }
                    }]
                }]
            }
        };
        styleSheet = new _StyleSheet(styleSheetNode);
    });

    describe("createStyle", () => {
        it("should clone an existing style", () => {
            const style = styleSheet.createStyle(0);
            expect(style.constructor).toBe(_Style);
            expect(styleSheet._node).toEqualJson({
                styleSheet: {
                    $: {
                        xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                    },
                    numFmts: [{
                        numFmt: []
                    }],
                    fonts: [{
                        $: {
                            'x14ac:knownFonts': 1
                        },
                        font: [
                            {}
                        ]
                    }],
                    fills: [{
                        $: {},
                        fill: [
                            {}
                        ]
                    }],
                    borders: [{
                        $: {},
                        border: [{
                            $: {
                                foo: "bar"
                            }
                        }, {
                            $: {
                                foo: "bar"
                            }
                        }]
                    }],
                    cellStyleXfs: [{
                        $: {
                            count: 1
                        },
                        xf: [{
                            $: {
                                numFmtId: 0,
                                fontId: 0,
                                fillId: 0,
                                borderId: 0
                            }
                        }]
                    }],
                    cellXfs: [{
                        $: {},
                        xf: [{
                            $: {
                                numFmtId: 0,
                                fontId: 0,
                                fillId: 0,
                                borderId: 0,
                                applyBorder: 1,
                                xfId: 0
                            }
                        }, {
                            $: {
                                numFmtId: 0,
                                fontId: 0,
                                fillId: 0,
                                borderId: 1,
                                applyFill: 1,
                                applyFont: 1,
                                applyBorder: 1,
                                xfId: 0
                            }
                        }]
                    }]
                }
            });
            expect(_Style).toHaveBeenCalledWith(
                styleSheet,
                1,
                styleSheetNode.styleSheet.cellXfs[0].xf[1],
                styleSheetNode.styleSheet.fonts[0].font[0],
                styleSheetNode.styleSheet.fills[0].fill[0],
                styleSheetNode.styleSheet.borders[0].border[1]
            );
        });

        it("should create a new style", () => {
            const style = styleSheet.createStyle(undefined);
            expect(style.constructor).toBe(_Style);
            expect(styleSheet._node).toEqualJson({
                styleSheet: {
                    $: {
                        xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                    },
                    numFmts: [{
                        numFmt: []
                    }],
                    fonts: [{
                        $: {
                            'x14ac:knownFonts': 1
                        },
                        font: [
                            {}
                        ]
                    }],
                    fills: [{
                        $: {},
                        fill: [
                            {}
                        ]
                    }],
                    borders: [{
                        $: {},
                        border: [{
                            $: {
                                foo: "bar"
                            }
                        }, {
                            left: [],
                            right: [],
                            top: [],
                            bottom: [],
                            diagonal: []
                        }]
                    }],
                    cellStyleXfs: [{
                        $: {
                            count: 1
                        },
                        xf: [{
                            $: {
                                numFmtId: 0,
                                fontId: 0,
                                fillId: 0,
                                borderId: 0
                            }
                        }]
                    }],
                    cellXfs: [{
                        $: {},
                        xf: [{
                            $: {
                                numFmtId: 0,
                                fontId: 0,
                                fillId: 0,
                                borderId: 0,
                                applyBorder: 1,
                                xfId: 0
                            }
                        }, {
                            $: {
                                fontId: 0,
                                fillId: 0,
                                borderId: 1,
                                applyFill: 1,
                                applyFont: 1,
                                applyBorder: 1
                            }
                        }]
                    }]
                }
            });
            expect(_Style).toHaveBeenCalledWith(
                styleSheet,
                1,
                styleSheetNode.styleSheet.cellXfs[0].xf[1],
                styleSheetNode.styleSheet.fonts[0].font[0],
                styleSheetNode.styleSheet.fills[0].fill[0],
                styleSheetNode.styleSheet.borders[0].border[1]
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
            expect(styleSheet._numFmtsNode).toEqualJson({ numFmt: [] });
            expect(styleSheet.getNumberFormatId('foo')).toBe(164);
            expect(styleSheet._numFmtsNode).toEqualJson({
                numFmt: [{
                    $: {
                        formatCode: "foo",
                        numFmtId: 164
                    }
                }]
            });
        });
    });

    describe("toObject", () => {
        it("should return the node as is", () => {
            expect(styleSheet.toObject()).toBe(styleSheetNode);
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

    describe("_initNode", () => {
        it("should add the numFmts node and clear the counts", () => {
            expect(styleSheetNode.styleSheet.numFmts).toEqualJson([{ numFmt: [] }]);
            expect(styleSheetNode.styleSheet.fonts[0].$.count).toBeUndefined();
            expect(styleSheetNode.styleSheet.fills[0].$.count).toBeUndefined();
            expect(styleSheetNode.styleSheet.borders[0].$.count).toBeUndefined();
            expect(styleSheetNode.styleSheet.cellXfs[0].$.count).toBeUndefined();
        });
    });
});
