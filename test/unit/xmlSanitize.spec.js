"use strict";

const proxyquire = require("proxyquire");

describe("xmlSanitize", () => {
    let xmlSanitize;

    beforeEach(() => {
        xmlSanitize = proxyquire("../../lib/xmlSanitize", {
            '@noCallThru': true
        });
    });

    it("should strip data link escape control char", () => {
        expect(xmlSanitize('test💯\x10content')).toEqual('test💯content');
    });

    it("should strip escape control char", () => {
        expect(xmlSanitize('test💯\x1Bcontent')).toEqual('test💯content');
    });

    it("should strip null control char", () => {
        expect(xmlSanitize('test💯\x00content')).toEqual('test💯content');
    });

    it("should strip unicode replacement char", () => {
        expect(xmlSanitize( 'Some �� Unicode characters')).toEqual('Some  Unicode characters');
    });

    it("should leave line breaks", () => {
        expect(xmlSanitize('test💯\ncontent')).toEqual('test💯\ncontent');
    });

    it("should leave carriage returns", () => {
        expect(xmlSanitize('test💯\r\ncontent')).toEqual('test💯\r\ncontent');
    });
});
