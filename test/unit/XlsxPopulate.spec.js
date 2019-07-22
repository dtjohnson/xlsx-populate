"use strict";

const proxyquire = require("proxyquire");

describe("XlsxPopulate", () => {
    let dateConverter, Workbook, XlsxPopulate, FormulaError;

    beforeEach(() => {
        dateConverter = jasmine.createSpyObj("dateConverter", ["dateToNumber", "numberToDate"]);
        dateConverter.dateToNumber.and.returnValue("NUMBER");
        dateConverter.numberToDate.and.returnValue("DATE");

        Workbook = jasmine.createSpyObj("Workbook", ["fromBlankAsync", "fromDataAsync", "fromFileAsync"]);
        Workbook.fromBlankAsync.and.returnValue("WORKBOOK");
        Workbook.fromDataAsync.and.returnValue("WORKBOOK");
        Workbook.fromFileAsync.and.returnValue("WORKBOOK");
        Workbook.MIME_TYPE = "MIME_TYPE";

        FormulaError = () => {};

        XlsxPopulate = proxyquire("../../src/XlsxPopulate", {
            './dateConverter': dateConverter,
            './Workbook': Workbook,
            './FormulaError': FormulaError,
            '@noCallThru': true
        });
    });

    describe("dateToNumber", () => {
        it("should call dateConverter.dateToNumber", () => {
            expect(XlsxPopulate.dateToNumber("DATE")).toBe("NUMBER");
            expect(dateConverter.dateToNumber).toHaveBeenCalledWith("DATE");
        });
    });

    describe("fromBlankAsync", () => {
        it("should call Workbook.fromBlankAsync", () => {
            expect(XlsxPopulate.fromBlankAsync()).toBe("WORKBOOK");
            expect(Workbook.fromBlankAsync).toHaveBeenCalledWith();
        });
    });

    describe("fromDataAsync", () => {
        it("should call Workbook.fromDataAsync", () => {
            expect(XlsxPopulate.fromDataAsync("DATA", "OPTS")).toBe("WORKBOOK");
            expect(Workbook.fromDataAsync).toHaveBeenCalledWith("DATA", "OPTS");
        });
    });

    describe("fromFileAsync", () => {
        it("should call Workbook.fromFileAsync", () => {
            expect(XlsxPopulate.fromFileAsync("PATH", "OPTS")).toBe("WORKBOOK");
            expect(Workbook.fromFileAsync).toHaveBeenCalledWith("PATH", "OPTS");
        });
    });

    describe("numberToDate", () => {
        it("should call dateConverter.numberToDate", () => {
            expect(XlsxPopulate.numberToDate("NUMBER")).toBe("DATE");
            expect(dateConverter.numberToDate).toHaveBeenCalledWith("NUMBER");
        });
    });

    describe("statics", () => {
        it("should set the statics", () => {
            expect(XlsxPopulate.MIME_TYPE).toBe("MIME_TYPE");
            expect(XlsxPopulate.FormulaError).toBe(FormulaError);
        });
    });
});
