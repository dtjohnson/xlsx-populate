"use strict";

const proxyquire = require("proxyquire").noCallThru();

describe("_ContentTypes", () => {
    let _ContentTypes, contentTypes, contentTypesNode;

    beforeEach(() => {
        _ContentTypes = proxyquire("../lib/_ContentTypes", {});

        contentTypesNode = {
            name: "Types",
            attributes: {
                xmlns: "http://schemas.openxmlformats.org/package/2006/content-types"
            },
            children: [
                {
                    name: "Default",
                    attributes: {
                        Extension: "bin",
                        ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"
                    }
                },
                {
                    name: "Override",
                    attributes: {
                        PartName: "/xl/workbook.xml",
                        ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
                    }
                },
                {
                    name: "Override",
                    attributes: {
                        PartName: "/xl/worksheets/sheet1.xml",
                        ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
                    }
                }
            ]
        };

        contentTypes = new _ContentTypes(contentTypesNode);
    });

    describe("add", () => {
        it("should add a new part", () => {
            contentTypes.add("NEW_PART_NAME", "NEW_CONTENT_TYPE");
            expect(contentTypesNode.children[3]).toEqualJson({
                name: "Override",
                attributes: {
                    PartName: "NEW_PART_NAME",
                    ContentType: "NEW_CONTENT_TYPE"
                }
            });
        });
    });

    describe("findByPartName", () => {
        it("should return the part if matched", () => {
            expect(contentTypes.findByPartName("/xl/worksheets/sheet1.xml")).toBe(contentTypesNode.children[2]);
            expect(contentTypes.findByPartName("/xl/workbook.xml")).toBe(contentTypesNode.children[1]);
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
