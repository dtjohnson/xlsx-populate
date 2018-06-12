"use strict";

const proxyquire = require("proxyquire");

describe("indexValidation", () => {
    let indexValidation;

    beforeEach(() => {
        indexValidation = proxyquire("../../lib/indexValidation", {
            '@noCallThru': true
        });
    });

    describe("validateIndex", () => {
        it("should do nothing for an index of 1", () => {
            indexValidation.validateIndex(1, "column");
        });

        it("should throw an exception for an index of 0", () => {
            expect(() => indexValidation.validateIndex(0, "column")).toThrow(
                new RangeError("Invaid column index 0. Remember that spreadsheets use 1 indexing.")
            );

            expect(() => indexValidation.validateIndex(0, "row")).toThrow(
                new RangeError("Invaid row index 0. Remember that spreadsheets use 1 indexing.")
            );
        });

        it("should throw an exception for an index of -1", () => {
            expect(() => indexValidation.validateIndex(-1, "column")).toThrow(
                new RangeError("Invaid column index -1. Remember that spreadsheets use 1 indexing.")
            );

            expect(() => indexValidation.validateIndex(-1, "row")).toThrow(
                new RangeError("Invaid row index -1. Remember that spreadsheets use 1 indexing.")
            );
        });
    });
});
