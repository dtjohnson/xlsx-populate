"use strict";

const proxyquire = require("proxyquire");

describe("regexify", () => {
    let regexify;

    beforeEach(() => {
        regexify = proxyquire("../../lib/regexify", {
            '@noCallThru': true
        });
    });

    it("should return the regex with lastIndex reset", () => {
        const regexp = /.+/;
        regexp.lastIndex = 5;

        const actual = regexify(regexp);
        expect(actual).toBe(regexp);
        expect(actual.lastIndex).toBe(0);
    });

    it("should convert a string to a regexp", () => {
        expect(regexify("search.[?")).toEqual(/search\.\[\?/gim);
    });
});
