"use strict";

const proxyquire = require("proxyquire");

describe("addressConverter", () => {
    let addressConverter;

    beforeEach(() => {
        addressConverter = proxyquire("../../lib/addressConverter", {
            '@noCallThru': true
        });
    });

    describe("columnNameToNumber", () => {
        it("should convert the name to a number", () => {
            expect(addressConverter.columnNameToNumber('A')).toBe(1);
            expect(addressConverter.columnNameToNumber('C')).toBe(3);
            expect(addressConverter.columnNameToNumber('Z')).toBe(26);
            expect(addressConverter.columnNameToNumber('AA')).toBe(27);
            expect(addressConverter.columnNameToNumber('ZZ')).toBe(702);
            expect(addressConverter.columnNameToNumber('AAC')).toBe(705);
        });
    });

    describe("columnNumberToName", () => {
        it("should convert the number to a name", () => {
            expect(addressConverter.columnNumberToName(1)).toBe('A');
            expect(addressConverter.columnNumberToName(3)).toBe('C');
            expect(addressConverter.columnNumberToName(26)).toBe('Z');
            expect(addressConverter.columnNumberToName(27)).toBe('AA');
            expect(addressConverter.columnNumberToName(702)).toBe('ZZ');
            expect(addressConverter.columnNumberToName(705)).toBe('AAC');
        });
    });

    describe("fromAddress", () => {
        it("should parse a range", () => {
            expect(addressConverter.fromAddress("A1:C3")).toEqualJson({
                type: 'range',
                startColumnAnchored: false,
                startColumnName: 'A',
                startColumnNumber: 1,
                startRowAnchored: false,
                startRowNumber: 1,
                endColumnAnchored: false,
                endColumnName: 'C',
                endColumnNumber: 3,
                endRowAnchored: false,
                endRowNumber: 3
            });

            expect(addressConverter.fromAddress("Sheet1!$B$4:$D$1")).toEqualJson({
                type: 'range',
                sheetName: 'Sheet1',
                startColumnAnchored: true,
                startColumnName: 'B',
                startColumnNumber: 2,
                startRowAnchored: true,
                startRowNumber: 4,
                endColumnAnchored: true,
                endColumnName: 'D',
                endColumnNumber: 4,
                endRowAnchored: true,
                endRowNumber: 1
            });
        });

        it("should parse a cell", () => {
            expect(addressConverter.fromAddress("Z56")).toEqualJson({
                type: 'cell',
                columnAnchored: false,
                columnName: 'Z',
                columnNumber: 26,
                rowAnchored: false,
                rowNumber: 56
            });

            expect(addressConverter.fromAddress("'Sheet One'!$AC$1")).toEqualJson({
                type: 'cell',
                sheetName: 'Sheet One',
                columnAnchored: true,
                columnName: 'AC',
                columnNumber: 29,
                rowAnchored: true,
                rowNumber: 1
            });
        });

        it("should parse a column range", () => {
            expect(addressConverter.fromAddress("Z:ZZ")).toEqualJson({
                type: 'columnRange',
                startColumnAnchored: false,
                startColumnName: 'Z',
                startColumnNumber: 26,
                endColumnAnchored: false,
                endColumnName: 'ZZ',
                endColumnNumber: 702
            });

            expect(addressConverter.fromAddress("'Foo''s Bar'!$A:$B")).toEqualJson({
                type: 'columnRange',
                sheetName: "Foo's Bar",
                startColumnAnchored: true,
                startColumnName: 'A',
                startColumnNumber: 1,
                endColumnAnchored: true,
                endColumnName: 'B',
                endColumnNumber: 2
            });
        });

        it("should parse a column", () => {
            expect(addressConverter.fromAddress("E:E")).toEqualJson({
                type: 'column',
                columnAnchored: false,
                columnName: 'E',
                columnNumber: 5
            });

            expect(addressConverter.fromAddress("'Foo!'!$A:$A")).toEqualJson({
                type: 'column',
                sheetName: "Foo!",
                columnAnchored: true,
                columnName: 'A',
                columnNumber: 1
            });
        });

        it("should parse a row range", () => {
            expect(addressConverter.fromAddress("103:104")).toEqualJson({
                type: 'rowRange',
                startRowAnchored: false,
                startRowNumber: 103,
                endRowAnchored: false,
                endRowNumber: 104
            });

            expect(addressConverter.fromAddress("Sheet1!$5:$3")).toEqualJson({
                type: 'rowRange',
                sheetName: 'Sheet1',
                startRowAnchored: true,
                startRowNumber: 5,
                endRowAnchored: true,
                endRowNumber: 3
            });
        });

        it("should parse a row", () => {
            expect(addressConverter.fromAddress("23:23")).toEqualJson({
                type: 'row',
                rowAnchored: false,
                rowNumber: 23
            });

            expect(addressConverter.fromAddress("Sheet1!$5:$5")).toEqualJson({
                type: 'row',
                sheetName: 'Sheet1',
                rowAnchored: true,
                rowNumber: 5
            });
        });

        it("should return undefined", () => {
            expect(addressConverter.fromAddress("Foo")).toBeUndefined();
        });
    });

    describe("toAddress", () => {
        it("should create a range address", () => {
            expect(addressConverter.toAddress({
                type: 'range',
                startColumnAnchored: false,
                startColumnName: 'A',
                startColumnNumber: 1,
                startRowAnchored: false,
                startRowNumber: 1,
                endColumnAnchored: false,
                endColumnName: 'C',
                endColumnNumber: 3,
                endRowAnchored: false,
                endRowNumber: 3
            })).toBe("A1:C3");

            expect(addressConverter.toAddress({
                type: 'range',
                sheetName: 'Sheet1',
                startColumnAnchored: true,
                startColumnName: 'B',
                startColumnNumber: 2,
                startRowAnchored: true,
                startRowNumber: 4,
                endColumnAnchored: true,
                endColumnName: 'D',
                endColumnNumber: 4,
                endRowAnchored: true,
                endRowNumber: 1
            })).toBe("'Sheet1'!$B$4:$D$1");
        });

        it("should create a cell address", () => {
            expect(addressConverter.toAddress({
                type: 'cell',
                columnAnchored: false,
                columnName: 'Z',
                columnNumber: 26,
                rowAnchored: false,
                rowNumber: 56
            })).toBe("Z56");

            expect(addressConverter.toAddress({
                type: 'cell',
                sheetName: 'Sheet One',
                columnAnchored: true,
                columnName: 'AC',
                columnNumber: 29,
                rowAnchored: true,
                rowNumber: 1
            })).toBe("'Sheet One'!$AC$1");
        });

        it("should create a column range address", () => {
            expect(addressConverter.toAddress({
                type: 'columnRange',
                startColumnAnchored: false,
                startColumnName: 'Z',
                startColumnNumber: 26,
                endColumnAnchored: false,
                endColumnName: 'ZZ',
                endColumnNumber: 702
            })).toBe("Z:ZZ");

            expect(addressConverter.toAddress({
                type: 'columnRange',
                sheetName: "Foo's Bar",
                startColumnAnchored: true,
                startColumnName: 'A',
                startColumnNumber: 1,
                endColumnAnchored: true,
                endColumnName: 'B',
                endColumnNumber: 2
            })).toBe("'Foo''s Bar'!$A:$B");
        });

        it("should create a column address", () => {
            expect(addressConverter.toAddress({
                type: 'column',
                columnAnchored: false,
                columnName: 'E',
                columnNumber: 5
            })).toBe("E:E");

            expect(addressConverter.toAddress({
                type: 'column',
                sheetName: "Foo!",
                columnAnchored: true,
                columnName: 'A',
                columnNumber: 1
            })).toBe("'Foo!'!$A:$A");
        });

        it("should create a row range address", () => {
            expect(addressConverter.toAddress({
                type: 'rowRange',
                startRowAnchored: false,
                startRowNumber: 103,
                endRowAnchored: false,
                endRowNumber: 104
            })).toBe("103:104");

            expect(addressConverter.toAddress({
                type: 'rowRange',
                sheetName: 'Sheet1',
                startRowAnchored: true,
                startRowNumber: 5,
                endRowAnchored: true,
                endRowNumber: 3
            })).toBe("'Sheet1'!$5:$3");
        });

        it("should create a row address", () => {
            expect(addressConverter.toAddress({
                type: 'row',
                rowAnchored: false,
                rowNumber: 23
            })).toBe("23:23");

            expect(addressConverter.toAddress({
                type: 'row',
                sheetName: 'Sheet1',
                rowAnchored: true,
                rowNumber: 5
            })).toBe("'Sheet1'!$5:$5");
        });
    });
});
