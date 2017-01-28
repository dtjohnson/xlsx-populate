"use strict";

/* eslint no-sync: "off" */

const fs = require("fs");

const data = fs.readFileSync(`${__dirname}/blank.xlsx`, "base64");
const template = fs.readFileSync(`${__dirname}/template.js`, "utf8");
const output = template.replace("{{DATA}}", data);
fs.writeFileSync(`${__dirname}/../lib/blank.js`, output);
