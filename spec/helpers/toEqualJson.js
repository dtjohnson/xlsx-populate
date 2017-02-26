"use strict";

const jsondiffpatch = require('jsondiffpatch');

beforeEach(() => {
    jasmine.addMatchers({
        toEqualJson(util, customEqualityTesters) {
            return {
                compare(actual, expected) {
                    const result = {};
                    actual = JSON.parse(JSON.stringify(actual));
                    expected = JSON.parse(JSON.stringify(expected));
                    result.pass = util.equals(actual, expected, customEqualityTesters);

                    if (!result.pass) {
                        result.name = "JSON objects don't match";
                        result.message = jsondiffpatch.formatters.console.format(jsondiffpatch.diff(expected, actual));
                    }

                    return result;
                }
            };
        }
    });
});
