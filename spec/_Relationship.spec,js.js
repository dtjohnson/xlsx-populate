"use strict";

const proxyquire = require("proxyquire").noCallThru();
const DOMParser = require('xmldom').DOMParser;
const parser = new DOMParser();

describe("_Relationship", () => {
    let _Relationship, relationship, relationshipNode;

    beforeEach(() => {
        _Relationship = proxyquire("../lib/_Relationship", {});
        relationshipNode = parser.parseFromString(`<Relationship Id="ID" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/TYPE" Target="TARGET"/>`).documentElement;
        relationship = new _Relationship(relationshipNode);
    });

    describe("constructor", () => {
        it("should store the node", () => {
            expect(relationship._node).toBe(relationshipNode);
        });
    });

    describe("id", () => {
        it("should return the ID", () => {
            expect(relationship.id()).toBe("ID");
        });
    });

    describe("type", () => {
        it("should return the type", () => {
            expect(relationship.type()).toBe("TYPE");
        });
    });
});
