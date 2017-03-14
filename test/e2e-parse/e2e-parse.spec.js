"use strict";

process.chdir(__dirname);

const fs = require("fs");
const glob = require("glob");
const path = require("path");

const XlsxPopulate = require("../../lib/XlsxPopulate");

// const testCases = ["./simple/"]; // To focus
const testCases = glob.sync("./*/");

describe("e2e-generate", () => {
    testCases.map(testCase => {
        itAsync(testCase, () => {
            return XlsxPopulate.fromFileAsync(`${testCase}input.xlsx`)
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
