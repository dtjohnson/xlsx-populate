"use strict";

const proxyquire = require("proxyquire");

describe("regexify", () => {
    let regexify;

    beforeEach(() => {
        regexify = proxyquire("../../lib/regexify", {
            '@noCallThru': true
        });
    });

    it("should return a regexp unchanged", () => {
        const regexp = /.+/;
        expect(regexify(regexp)).toBe(regexp);
    });

    it("should convert a string to a regexp", () => {
        expect(regexify("search.[?")).toEqual(/search\.\[\?/gim);
    });
});
