"use strict";

const proxyquire = require("proxyquire").noCallThru();
const DOMParser = require('xmldom').DOMParser;
const parser = new DOMParser();

describe("_ContentType", () => {
    let _ContentType, contentType, contentTypeNode;

    beforeEach(() => {
        _ContentType = proxyquire("../lib/_ContentType", {});
        contentTypeNode = parser.parseFromString('<Override PartName="PART_NAME" ContentType="CONTENT_TYPE"/>').documentElement;
        contentType = new _ContentType(contentTypeNode);
    });

    describe("constructor", () => {
        it("should store the node", () => {
            expect(contentType._node).toBe(contentTypeNode);
        });
    });

    describe("partName", () => {
        it("should return the part name", () => {
            expect(contentType.partName()).toBe("PART_NAME");
        });
    });

    describe("contentType", () => {
        it("should return the content type", () => {
            expect(contentType.contentType()).toBe("CONTENT_TYPE");
        });
    });
});
