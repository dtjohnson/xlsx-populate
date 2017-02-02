"use strict";

const proxyquire = require("proxyquire").noCallThru();

describe("_SharedStrings", () => {
    let _SharedStrings, sharedStrings, sharedStringsNode;

    beforeEach(() => {
        _SharedStrings = proxyquire("../lib/_SharedStrings", {});

        sharedStringsNode = {
            sst: {
                $: {
                    xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
                    count: 3,
                    unique: 7
                },
                si: [
                    { t: ["foo"] }
                ]
            }
        };

        sharedStrings = new _SharedStrings(sharedStringsNode);
    });

    describe("getIndexForString", () => {
        beforeEach(() => {
            sharedStrings._stringArray = ["foo", "bar"];
            sharedStrings._indexMap = { foo: 0, bar: 1 };
        });

        it("should return the index if the string already exists", () => {
            expect(sharedStrings.getIndexForString("foo")).toBe(0);
            expect(sharedStrings.getIndexForString("bar")).toBe(1);
        });

        it("should create a new entry if the string doesn't exist", () => {
            expect(sharedStrings.getIndexForString("baz")).toBe(2);
            expect(sharedStrings._stringArray).toEqual(["foo", "bar", "baz"]);
            expect(sharedStrings._indexMap).toEqual({ foo: 0, bar: 1, baz: 2 });
            expect(sharedStringsNode.sst.si[sharedStringsNode.sst.si.length - 1]).toEqualJson({
                t: ["baz"]
            });
        });
    });

    describe("getStringByIndex", () => {
        it("should return the string at a given index", () => {
            sharedStrings._stringArray = ["foo", "bar", "baz"];
            expect(sharedStrings.getStringByIndex(0)).toBe("foo");
            expect(sharedStrings.getStringByIndex(1)).toBe("bar");
            expect(sharedStrings.getStringByIndex(2)).toBe("baz");
            expect(sharedStrings.getStringByIndex(3)).toBeUndefined();
        });
    });

    describe("toObject", () => {
        it("should return the node as is", () => {
            expect(sharedStrings.toObject()).toBe(sharedStringsNode);
        });
    });

    describe("_cacheExistingSharedStrings", () => {
        it("should cache the existing shared strings", () => {
            sharedStrings._siNode = [
                { t: ["foo"] },
                { t: ["bar"] },
                { r: [{}] },
                { t: ["baz"] }
            ];

            sharedStrings._stringArray = [];
            sharedStrings._indexMap = {};
            sharedStrings._cacheExistingSharedStrings();

            expect(sharedStrings._stringArray).toEqualJson([
                "foo",
                "bar",
                { r: [{}] },
                "baz"
            ]);
            expect(sharedStrings._indexMap).toEqualJson({
                foo: 0,
                bar: 1,
                baz: 3
            });
        });
    });

    describe("_initNode", () => {
        it("should create the node if needed", () => {
            sharedStrings._initNode(null);
            expect(sharedStrings._node).toEqualJson({
                sst: {
                    $: {
                        xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                    },
                    si: []
                }
            });
        });

        it("should set the _siNode and clear the counts", () => {
            expect(sharedStrings._siNode).toEqualJson([
                { t: ["foo"] }
            ]);
            expect(sharedStrings._node).toEqualJson({
                sst: {
                    $: {
                        xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                    },
                    si: [
                        { t: ["foo"] }
                    ]
                }
            });
        });
    });
});
