"use strict";

process.chdir(__dirname);

const fs = require("fs");
const glob = require("glob");
const path = require("path");

const XlsxPopulate = require("../../lib/XlsxPopulate");

// const testCases = ["./encrypted/"]; // To focus
const testCases = glob.sync("./*/");

describe("e2e-parse", () => {
    testCases.map(testCase => {
        itAsync(testCase, () => {
            const password = fs.existsSync(`${testCase}password.txt`) && fs.readFileSync(`${testCase}password.txt`, "utf8");
            return XlsxPopulate.fromFileAsync(`${testCase}input.xlsx`, { password })
                .then(workbook => {
                    const parse = require(`${testCase}parse`);
                    return parse(workbook);
                })
                .then(results => {
                    const expected = JSON.parse(fs.readFileSync(`${testCase}expected.json`));
                    expect(results).toEqualJson(expected);
                });
        });
    });
});
