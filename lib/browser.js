"use strict";

//const debug = require("./debug")('browser');

console.log("hit2");

const fs = require("fs");
const wb = fs.readFileSync(__dirname + "/blank.xlsx");

module.exports = () => {
    console.log(wb);
    throw new Error("bad");
};
