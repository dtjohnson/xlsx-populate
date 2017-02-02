"use strict";

const proxyquire = require("proxyquire").noCallThru();

describe("_ContentTypes", () => {
    let _ContentTypes, contentTypes, contentTypesNode;

    beforeEach(() => {
        _ContentTypes = proxyquire("../lib/_ContentTypes", {});

        contentTypesNode = {
            Types: {
                $: {
                    xmlns: "http://schemas.openxmlformats.org/package/2006/content-types"
                },
                Default: [
                    {
                        $: {
                            Extension: "rels",
                            ContentType: "application/vnd.openxmlformats-package.relationships+xml"
                        }
                    }
                ],
                Override: [
                    {
                        $: {
                            PartName: "/xl/workbook.xml",
                            ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml                            "
                        }
                    },
                    {
                        $: {
                            PartName: "/xl/worksheets/sheet1.xml",
                            ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"

                        }
                    }
                ]
            }
        };

        contentTypes = new _ContentTypes(contentTypesNode);
    });

    describe("add", () => {
        it("should add a new part", () => {
            contentTypes.add("NEW_PART_NAME", "NEW_CONTENT_TYPE");
            expect(contentTypesNode.Types.Override[2]).toEqualJson({
                $: {
                    PartName: "NEW_PART_NAME",
                    ContentType: "NEW_CONTENT_TYPE"
                }
            });
        });
    });

    describe("findByPartName", () => {
        it("should return the part if matched", () => {
            expect(contentTypes.findByPartName("/xl/worksheets/sheet1.xml")).toBe(contentTypesNode.Types.Override[1]);
            expect(contentTypes.findByPartName("/xl/workbook.xml")).toBe(contentTypesNode.Types.Override[0]);
        });

        it("should return undefined if not matched", () => {
            expect(contentTypes.findByPartName("foo")).toBeUndefined();
        });
    });

    describe("toObject", () => {
        it("should return the node as is", () => {
            expect(contentTypes.toObject()).toBe(contentTypesNode);
        });
    });
});
