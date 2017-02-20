"use strict";

const escapeStringRegexp = require('escape-string-regexp');

/**
 * Convert a pattern to a RegExp.
 * @param {RegExp|string} pattern - The pattern to convert.
 * @returns {RegExp} The regex.
 * @private
 */
module.exports = pattern => {
    return typeof pattern === "string" ? new RegExp(escapeStringRegexp(pattern), "igm") : pattern;
};
