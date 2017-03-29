"use strict";

const proxyquire = require("proxyquire");

describe("dateConverter", () => {
    let dateConverter;

    beforeEach(() => {
        dateConverter = proxyquire("../../lib/dateConverter", {
            '@noCallThru': true
        });
    });

    describe("dateToNumber", () => {
        it("should convert date to number", () => {
            expect(dateConverter.dateToNumber(new Date('01 Jan 1900 00:00:00'))).toBe(1);
            expect(dateConverter.dateToNumber(new Date('28 Feb 1900 00:00:00'))).toBe(59);
            expect(dateConverter.dateToNumber(new Date('01 Mar 1900 00:00:00'))).toBe(61);
            expect(dateConverter.dateToNumber(new Date('07 Mar 2015 13:26:24'))).toBe(42070.56);
            expect(dateConverter.dateToNumber(new Date('04 Apr 2017 20:00:00'))).toBeCloseTo(42829.8333333333, 10);
        });
    });

    describe("numberToDate", () => {
        it("should convert number to date", () => {
            expect(dateConverter.numberToDate(1)).toEqual(new Date('01 Jan 1900 00:00:00'));
            expect(dateConverter.numberToDate(59)).toEqual(new Date('28 Feb 1900 00:00:00'));
            expect(dateConverter.numberToDate(61)).toEqual(new Date('01 Mar 1900 00:00:00'));
            expect(dateConverter.numberToDate(42070.56)).toEqual(new Date('07 Mar 2015 13:26:24'));
            expect(dateConverter.numberToDate(42829.8333333333)).toEqual(new Date('04 Apr 2017 20:00:00'));
        });
    });
});
