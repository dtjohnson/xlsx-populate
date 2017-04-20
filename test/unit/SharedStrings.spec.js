"use strict";

const proxyquire = require("proxyquire");

describe("SharedStrings", () => {
    let SharedStrings, sharedStrings, sharedStringsNode;

    beforeEach(() => {
        SharedStrings = proxyquire("../../lib/SharedStrings", {
            '@noCallThru': true
        });

        sharedStringsNode = {
            name: "sst",
            attributes: {
                xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
                count: 3,
                uniqueCount: 7
            },
            children: [
                {
                    name: "si",
                    children: [
                        {
                            name: "t",
                            children: ["foo"]
                        }
                    ]
                }
            ]
        };

        sharedStrings = new SharedStrings(sharedStringsNode);
    });

    describe("getIndexForString", () => {
        beforeEach(() => {
            sharedStrings._stringArray = [
                "foo",
                "bar",
                [{ name: "r", children: [{}] }, { name: "r", children: [{}] }]
            ];

            sharedStrings._indexMap = {
                foo: 0,
                bar: 1,
                '[{"name":"r","children":[{}]},{"name":"r","children":[{}]}]': 2
            };
        });

        it("should return the index if the string already exists", () => {
            expect(sharedStrings.getIndexForString("foo")).toBe(0);
            expect(sharedStrings.getIndexForString("bar")).toBe(1);
            expect(sharedStrings.getIndexForString([{ name: "r", children: [{}] }, { name: "r", children: [{}] }])).toBe(2);
        });

        it("should create a new entry if the string doesn't exist", () => {
            expect(sharedStrings.getIndexForString("baz")).toBe(3);
            expect(sharedStrings._stringArray).toEqualJson([
                "foo",
                "bar",
                [{ name: "r", children: [{}] }, { name: "r", children: [{}] }],
                "baz"
            ]);
            expect(sharedStrings._indexMap).toEqualJson({
                foo: 0,
                bar: 1,
                '[{"name":"r","children":[{}]},{"name":"r","children":[{}]}]': 2,
                baz: 3
            });
            expect(sharedStringsNode.children[sharedStringsNode.children.length - 1]).toEqualJson({
                name: "si",
                children: [
                    {
                        name: "t",
                        attributes: { 'xml:space': "preserve" },
                        children: ["baz"]
                    }
                ]
            });
        });

        it("should create a new array entry if the array doesn't exist", () => {
            expect(sharedStrings.getIndexForString([{ name: "r", children: [{}] }])).toBe(3);
            expect(sharedStrings._stringArray).toEqualJson([
                "foo",
                "bar",
                [{ name: "r", children: [{}] }, { name: "r", children: [{}] }],
                [{ name: "r", children: [{}] }]
            ]);
            expect(sharedStrings._indexMap).toEqualJson({
                foo: 0,
                bar: 1,
                '[{"name":"r","children":[{}]},{"name":"r","children":[{}]}]': 2,
                '[{"name":"r","children":[{}]}]': 3
            });
            expect(sharedStringsNode.children[sharedStringsNode.children.length - 1]).toEqualJson({
                name: "si",
                children: [{ name: "r", children: [{}] }]
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

    describe("toXml", () => {
        it("should return the node as is", () => {
            expect(sharedStrings.toXml()).toBe(sharedStringsNode);
        });
    });

    describe("_cacheExistingSharedStrings", () => {
        it("should cache the existing shared strings", () => {
            sharedStrings._node.children = [
                { name: "si", children: [{ name: "t", children: ["foo"] }] },
                { name: "si", children: [{ name: "t", children: ["bar"] }] },
                { name: "si", children: [{ name: "r", children: [{}] }, { name: "r", children: [{}] }] },
                { name: "si", children: [{ name: "t", children: ["baz"] }] }
            ];

            sharedStrings._stringArray = [];
            sharedStrings._indexMap = {};
            sharedStrings._cacheExistingSharedStrings();

            expect(sharedStrings._stringArray).toEqualJson([
                "foo",
                "bar",
                [{ name: "r", children: [{}] }, { name: "r", children: [{}] }],
                "baz"
            ]);
            expect(sharedStrings._indexMap).toEqualJson({
                foo: 0,
                bar: 1,
                '[{"name":"r","children":[{}]},{"name":"r","children":[{}]}]': 2,
                baz: 3
            });
        });
    });

    describe("_init", () => {
        it("should create the node if needed", () => {
            sharedStrings._init(null);
            expect(sharedStrings._node).toEqualJson({
                name: "sst",
                attributes: {
                    xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                },
                children: []
            });
        });

        it("should clear the counts", () => {
            expect(sharedStrings._node.attributes).toEqualJson({
                xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            });
        });
    });
});
