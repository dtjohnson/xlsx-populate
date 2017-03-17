"use strict";

const proxyquire = require("proxyquire");

describe("FormulaError", () => {
    let FormulaError;

    beforeEach(() => {
        FormulaError = proxyquire("../../lib/FormulaError", {
            '@noCallThru': true
        });
    });

    describe("error", () => {
        it("should return the error", () => {
            const formulaError = new FormulaError("foo");
            expect(formulaError.error()).toBe("foo");
        });
    });

    describe("static", () => {
        it("should create the static instances", () => {
            expect(FormulaError.DIV0.error()).toBe("#DIV/0!");
            expect(FormulaError.NA.error()).toBe("#N/A");
            expect(FormulaError.NAME.error()).toBe("#NAME?");
            expect(FormulaError.NULL.error()).toBe("#NULL!");
            expect(FormulaError.NUM.error()).toBe("#NUM!");
            expect(FormulaError.REF.error()).toBe("#REF!");
            expect(FormulaError.VALUE.error()).toBe("#VALUE!");
        });
    });

    describe("getError", () => {
        it("should get the matching error", () => {
            expect(FormulaError.getError("#VALUE!")).toBe(FormulaError.VALUE);
        });

        it("should create a new instance for unknown errors", () => {
            expect(FormulaError.getError("foo").error()).toBe("foo");
        });
    });
});
