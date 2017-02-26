"use strict";

// Export as a function as proxyquireify has trouble with constant exports.
module.exports = () => new Buffer("{{DATA}}", "base64");
