"use strict";

const proxyquire = require("proxyquire");
const Promise = require("jszip").external.Promise;

describe("XlsxPopulate", () => {
    let dateConverter, Workbook, XlsxPopulate, FormulaError, externals;

    beforeEach(() => {
        dateConverter = jasmine.createSpyObj("dateConverter", ["dateToNumber", "numberToDate"]);
        dateConverter.dateToNumber.and.returnValue("NUMBER");
        dateConverter.numberToDate.and.returnValue("DATE");

        Workbook = jasmine.createSpyObj("Workbook", ["fromBlankAsync", "fromDataAsync", "fromFileAsync"]);
        Workbook.fromBlankAsync.and.returnValue("WORKBOOK");
        Workbook.fromDataAsync.and.returnValue("WORKBOOK");
        Workbook.fromFileAsync.and.returnValue("WORKBOOK");
        Workbook.MIME_TYPE = "MIME_TYPE";

        // proxyquire doesn't like overriding raw objects... a spy obj works.
        externals = jasmine.createSpyObj("externals", ["_"]);
        externals.Promise = Promise;

        FormulaError = () => {};

        XlsxPopulate = proxyquire("../../lib/XlsxPopulate", {
            './externals': externals,
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

    describe("Promise", () => {
        it("should get/set the Promise", () => {
            expect(XlsxPopulate.Promise).toBeDefined();
            expect(XlsxPopulate.Promise.all).toEqual(jasmine.any(Function));
            XlsxPopulate.Promise = "PROMISE";
            expect(XlsxPopulate.Promise).toBe("PROMISE");
        });
    });

    describe("statics", () => {
        it("should set the statics", () => {
            expect(XlsxPopulate.MIME_TYPE).toBe("MIME_TYPE");
            expect(XlsxPopulate.FormulaError).toBe(FormulaError);
        });
    });
});
