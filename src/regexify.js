"use strict";

const _ = require("lodash");

/**
 * Convert a pattern to a RegExp.
 * @param {RegExp|string} pattern - The pattern to convert.
 * @returns {RegExp} The regex.
 * @private
 */
module.exports = pattern => {
    if (typeof pattern === "string") {
        pattern = new RegExp(_.escapeRegExp(pattern), "igm");
    }

    pattern.lastIndex = 0;

    return pattern;
};
