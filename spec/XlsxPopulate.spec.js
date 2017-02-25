"use strict";

const proxyquire = require("proxyquire").noCallThru();

describe("XlsxPopulate", () => {
    let dateConverter, Workbook, XlsxPopulate;

    beforeEach(() => {
        dateConverter = jasmine.createSpyObj("dateConverter", ["dateToNumber", "numberToDate"]);
        dateConverter.dateToNumber.and.returnValue("NUMBER");
        dateConverter.numberToDate.and.returnValue("DATE");

        Workbook = jasmine.createSpyObj("Workbook", ["fromBlankAsync", "fromDataAsync", "fromFileAsync"]);
        Workbook.fromBlankAsync.and.returnValue("WORKBOOK");
        Workbook.fromDataAsync.and.returnValue("WORKBOOK");
        Workbook.fromFileAsync.and.returnValue("WORKBOOK");

        XlsxPopulate = proxyquire("../lib/XlsxPopulate", {
            './dateConverter': dateConverter,
            './Workbook': Workbook
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
            expect(XlsxPopulate.fromDataAsync("DATA")).toBe("WORKBOOK");
            expect(Workbook.fromDataAsync).toHaveBeenCalledWith("DATA");
        });
    });

    describe("fromFileAsync", () => {
        it("should call Workbook.fromFileAsync", () => {
            expect(XlsxPopulate.fromFileAsync("PATH")).toBe("WORKBOOK");
            expect(Workbook.fromFileAsync).toHaveBeenCalledWith("PATH");
        });
    });

    describe("numberToDate", () => {
        it("should call dateConverter.numberToDate", () => {
            expect(XlsxPopulate.numberToDate("NUMBER")).toBe("DATE");
            expect(dateConverter.numberToDate).toHaveBeenCalledWith("NUMBER");
        });
    });
});
