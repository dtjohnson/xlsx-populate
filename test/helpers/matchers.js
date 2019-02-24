"use strict";

beforeEach(() => {
    jasmine.addMatchers({
        toEqualUInt8Array() {
            return {
                compare(actual, expected) {
                    const result = { pass: true };

                    if (actual.byteLength === expected.byteLength) {
                        for (let i = 0; i < actual.byteLength; i++) {
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
            };
        }
    });
});
