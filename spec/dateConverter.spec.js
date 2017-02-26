"use strict";

const proxyquire = require("proxyquire");

describe("dateConverter", () => {
    let dateConverter;

    beforeEach(() => {
        dateConverter = proxyquire("../lib/dateConverter", {
            '@noCallThru': true
        });
    });

    describe("dateToNumber", () => {
        it("should convert date to number", () => {
            expect(dateConverter.dateToNumber(new Date('01 Jan 1900 00:00:00'))).toBe(1);
            expect(dateConverter.dateToNumber(new Date('28 Feb 1900 00:00:00'))).toBe(59);
            expect(dateConverter.dateToNumber(new Date('01 Mar 1900 00:00:00'))).toBe(61);
            expect(dateConverter.dateToNumber(new Date('07 Mar 2015 13:26:24'))).toBe(42070.56);
        });
    });

    describe("unumberToDate", () => {
        it("should convert number to date", () => {
            expect(dateConverter.numberToDate(1)).toEqual(new Date('01 Jan 1900 00:00:00'));
            expect(dateConverter.numberToDate(59)).toEqual(new Date('28 Feb 1900 00:00:00'));
            expect(dateConverter.numberToDate(61)).toEqual(new Date('01 Mar 1900 00:00:00'));
            expect(dateConverter.numberToDate(42070.56)).toEqual(new Date('07 Mar 2015 13:26:24'));
        });
    });
});
