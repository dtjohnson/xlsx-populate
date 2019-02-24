"use strict";

module.exports = config => {
    config.set({
        // logLevel: config.LOG_DEBUG,
        frameworks: ["jasmine", "karma-typescript"],
        files: [
            "src/**/*.ts"
        ],
        preprocessors: {
            "**/*.ts": "karma-typescript"
        },
        reporters: ["progress", "karma-typescript"],
        browsers: ["Chrome"],
        karmaTypescriptConfig: {
            include: ["src"],
            exclude: [],
            tsconfig: "./tsconfig.karma.json"
        }
    });
};
