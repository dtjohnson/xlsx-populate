"use strict";

const proxyquire = require("proxyquire").noCallThru();

fdescribe("_Relationships", () => {
    let _Relationships, relationships, relationshipsNode;

    beforeEach(() => {
        _Relationships = proxyquire("../lib/_Relationships", {});

        relationshipsNode = {
            Relationships: {
                $: {
                    xmlns: "http://schemas.openxmlformats.org/package/2006/relationships"
                },
                Relationship: [
                    {
                        $: {
                            Id: "rId2",
                            Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
                            Target: "theme/theme1.xml"
                        }
                    },
                    {
                        $: {
                            Id: "rId1",
                            Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                            Target: "worksheets/sheet1.xml"
                        }
                    }
                ]
            }
        };

        relationships = new _Relationships(relationshipsNode);
    });

    describe("add", () => {
        it("should add a new relationship", () => {
            spyOn(Date, "now").and.returnValue(12345);
            relationships.add("TYPE", "TARGET");
            expect(relationshipsNode.Relationships.Relationship[2]).toEqualJson({
                $: {
                    Id: "rId12345",
                    Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/TYPE",
                    Target: "TARGET"
                }
            });
        });
    });

    describe("findByType", () => {
        it("should return the relationship if matched", () => {
            expect(relationships.findByType("worksheet")).toBe(relationshipsNode.Relationships.Relationship[1])
            expect(relationships.findByType("theme")).toBe(relationshipsNode.Relationships.Relationship[0])
        });

        it("should return undefined if not matched", () => {
            expect(relationships.findByType("foo")).toBeUndefined();
        });
    });

    describe("toObject", () => {
        it("should return the node as is", () => {
            expect(relationships.toObject()).toBe(relationshipsNode);
        });
    });
});
