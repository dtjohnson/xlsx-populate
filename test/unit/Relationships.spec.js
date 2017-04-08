"use strict";

const proxyquire = require("proxyquire");

describe("Relationships", () => {
    let Relationships, relationships, relationshipsNode;

    beforeEach(() => {
        Relationships = proxyquire("../../lib/Relationships", {
            '@noCallThru': true
        });

        relationshipsNode = {
            name: "Relationships",
            attributes: {
                xmlns: "http://schemas.openxmlformats.org/package/2006/relationships"
            },
            children: [
                {
                    name: "Relationship",
                    attributes: {
                        Id: "rId2",
                        Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
                        Target: "theme/theme1.xml"
                    }
                },
                {
                    name: "Relationship",
                    attributes: {
                        Id: "rId1",
                        Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                        Target: "worksheets/sheet1.xml"
                    }
                }
            ]
        };

        relationships = new Relationships(relationshipsNode);
    });

    describe("add", () => {
        it("should add a new relationship", () => {
            relationships.add("TYPE", "TARGET");
            expect(relationshipsNode.children[2]).toEqualJson({
                name: "Relationship",
                attributes: {
                    Id: "rId3",
                    Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/TYPE",
                    Target: "TARGET"
                }
            });
        });

        it("should add a new relationship with target mode", () => {
            relationships.add("TYPE", "TARGET", "TARGET_MODE");
            expect(relationshipsNode.children[2]).toEqualJson({
                name: "Relationship",
                attributes: {
                    Id: "rId3",
                    Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/TYPE",
                    Target: "TARGET",
                    TargetMode: "TARGET_MODE"
                }
            });
        });
    });

    describe("findById", () => {
        it("should return the relationship if matched", () => {
            expect(relationships.findById("rId1")).toBe(relationshipsNode.children[1]);
            expect(relationships.findById("rId2")).toBe(relationshipsNode.children[0]);
        });

        it("should return undefined if not matched", () => {
            expect(relationships.findById("rId5")).toBeUndefined();
        });
    });

    describe("findByType", () => {
        it("should return the relationship if matched", () => {
            expect(relationships.findByType("worksheet")).toBe(relationshipsNode.children[1]);
            expect(relationships.findByType("theme")).toBe(relationshipsNode.children[0]);
        });

        it("should return undefined if not matched", () => {
            expect(relationships.findByType("foo")).toBeUndefined();
        });
    });

    describe("toXml", () => {
        it("should return the node as is", () => {
            expect(relationships.toXml()).toBe(relationshipsNode);
        });

        it("should return undefined", () => {
            relationshipsNode.children.length = 0;
            expect(relationships.toXml()).toBeUndefined();
        });
    });

    describe("_getStartingId", () => {
        it("should set the next ID to 1 if no children", () => {
            relationships._node.children = [];
            relationships._getStartingId();
            expect(relationships._nextId).toBe(1);
        });

        it("should set the next ID to last found ID + 1", () => {
            relationships._node.children = [
                { attributes: { Id: 'rId2' } },
                { attributes: { Id: 'rId1' } },
                { attributes: { Id: 'rId3' } }
            ];
            relationships._getStartingId();
            expect(relationships._nextId).toBe(4);
        });
    });

    describe("_init", () => {
        it("should create the node if needed", () => {
            relationships._init(null);
            expect(relationships._node).toEqualJson({
                name: "Relationships",
                attributes: {
                    xmlns: "http://schemas.openxmlformats.org/package/2006/relationships"
                },
                children: []
            });
        });
    });
});
