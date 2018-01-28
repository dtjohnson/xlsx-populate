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
        },

        toEqualUInt8Array() {
            return {
                compare(actual, expected) {
                    const result = { pass: true };

                    if (actual.byteLength === expected.byteLength) {
                        for (var i = 0; i < actual.byteLength; i++) {
                            if (actual[i] !== expected[i]) {
                                result.pass = false;
                                break;
                            }
                        }
                    } else {
                        result.pass = false;
                    }

                    if (!result.pass) result.name = "UInt8Arrays do not match";

                    return result;
                }
            }
        }
    });
});
