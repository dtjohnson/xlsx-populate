"use strict";

module.exports = function(config) {
    config.set({

        frameworks: ["jasmine", "karma-typescript"],

        files: [
            { pattern: "src/**/*.ts" },
            { pattern: "test/unit/**/*.ts" }
        ],

        preprocessors: {
            "src/**/*.ts": ["karma-typescript", "coverage"],
            "test/unit/**/*.ts": ["karma-typescript"]
        },

        reporters: ["progress", "coverage", "karma-typescript"],

        browsers: ["Chrome"],

        karmaTypescriptConfig: {
            tsconfig: "./test/tsconfig.json",
        },
    });
};
