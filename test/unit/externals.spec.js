"use strict";

const proxyquire = require("proxyquire");

describe("externals", () => {
    let externals, JSZip;

    beforeEach(() => {
        JSZip = jasmine.createSpyObj("JSZip", ["_"]);
        JSZip.external = { Promise: "PROMISE" };

        externals = proxyquire("../../lib/externals", {
            jszip: JSZip,
            '@noCallThru': true
        });
    });

    describe("Promise", () => {
        it("should get/set the JSZip Promise", () => {
            expect(externals.Promise).toBe("PROMISE");
            externals.Promise = "NEW PROMISE";
            expect(externals.Promise).toBe("NEW PROMISE");
        });
    });
});
