"use strict";

//const debug = require("./debug")('browser');

console.log("hit");

module.exports = () => {
    console.log("hit");
    throw new Error("bad");
};
