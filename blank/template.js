"use strict";

// Export as a function as proxyquireify has trouble with constant exports.
module.exports = () => Buffer.from("{{DATA}}", "base64");
