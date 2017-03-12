"use strict";

const glob = require("glob");
const edge = require('edge');

module.exports = path => new Promise((resolve, reject) => {
    // Install VSTO redistributable from here: https://www.microsoft.com/en-us/download/details.aspx?id=48217
    const interopPath = glob.sync("C:\\Windows\\assembly\\GAC_MSIL\\Microsoft.Office.Interop.Excel\\*\\Microsoft.Office.Interop.Excel.dll")[0];

    const helloWorld = edge.func({
        source: path,
        references: [interopPath]
    });

    helloWorld(3, (err, result) => {
        if (err) return reject(err);
        resolve(result);
    });
});
