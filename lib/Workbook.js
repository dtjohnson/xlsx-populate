"use strict";

var fs = require('fs');
var JSZip = require('jszip');
var utils = require('./utils');
var Sheet = require('./Sheet');
var path = require("path");
var DOMParser = require('xmldom').DOMParser;
var parser = new DOMParser();
var xpath = require("./xpath");

/**
 * Initializes a new Workbook.
 * @param {Buffer} data - File buffer of the Excel workbook.
 * @constructor
 */
var Workbook = function (data) {
    this._initialize(data);
};

/**
 * Initialize the workbook. (This is separated from the constructor to ease testing.)
 * @param {Buffer} data - File buffer of the Excel workbook.
 * @returns {undefined}
 * @private
 */
Workbook.prototype._initialize = function (data) {
    this._zip = new JSZip(data, { base64: false, checkCRC32: true });
    var workbookText = this._zip.file("xl/workbook.xml").asText();
    this._workbookXML = parser.parseFromString(workbookText).documentElement;

    this._sheets = [];
    var sheetNodes = xpath("sml:sheets/sml:sheet", this._workbookXML);

    for (var i = 0; i < sheetNodes.length; i++) {
        var sheetText = this._zip.file("xl/worksheets/sheet" + (i + 1) + ".xml").asText();
        var sheetXML = parser.parseFromString(sheetText).documentElement;

        // This is a blunt way to make sure formula values get updated.
        // It just clears all stored values in case the referenced cell values change.
        var valueNodes = xpath("sml:sheetData/sml:row/sml:c/sml:f/../sml:*[name(.) !='f']", sheetXML);
        valueNodes.forEach(function (valueNode) {
            valueNode.parentNode.removeChild(valueNode);
        });

        var sheet = new Sheet(this, sheetNodes[i], sheetXML);
        this._sheets.push(sheet);
    }
};

/**
 * Gets the sheet with the provided name or index (0-based).
 * @param {string|number} sheetNameOrIndex - The sheet name or index.
 * @returns {Sheet} The sheet, if found.
 */
Workbook.prototype.getSheet = function (sheetNameOrIndex) {
    if (utils.isInteger(sheetNameOrIndex)) return this._sheets[sheetNameOrIndex];

    for (var i = 0; i < this._sheets.length; i++) {
        var sheet = this._sheets[i];
        if (sheet.getName() === sheetNameOrIndex) return sheet;
    }
};

/**
 * Get a named cell. (Assumes names with workbook scope pointing to single cells.)
 * @param {string} cellName - The name of the cell.
 * @returns {Cell} The cell, if found.
 */
Workbook.prototype.getNamedCell = function (cellName) {
    var definedName = xpath("sml:definedNames/sml:definedName[@name='" + cellName + "']", this._workbookXML)[0];
    if (!definedName) return;

    var address = definedName.firstChild.nodeValue;
    var ref = utils.addressToRowAndColumn(address);
    if (!ref) return;

    return this.getSheet(ref.sheet).getCell(ref.row, ref.column);
};

/**
 * Gets the output.
 * @returns {Buffer} A node buffer for the generated Excel workbook.
 */
Workbook.prototype.output = function () {
    this._zip.file("xl/workbook.xml", this._workbookXML.toString());

    for (var i = 0; i < this._sheets.length; i++) {
        var index = i + 1;
        var sheet = this._sheets[i];
        this._zip.file("xl/worksheets/sheet" + index + ".xml", sheet._sheetXML.toString());
    }

    // Kill the calc chain since this will corrupt the file is formulas are removed.
    this._zip.remove("xl/calcChain.xml");

    return this._zip.generate({ type: "nodebuffer" });
};

/**
 * Writes to file with the given path.
 * @param {string} path - The path of the file.
 * @param {function} cb - A callback.
 * @returns {undefined}
 */
Workbook.prototype.toFile = function (path, cb) {
    fs.writeFile(path, this.output(), cb);
};

/**
 * Writes to file with the given path synchronously.
 * @param {string} path - The path of the file.
 * @returns {undefined}
 */
Workbook.prototype.toFileSync = function (path) {
    fs.writeFileSync(path, this.output());
};

/**
 * Creates a Workbook from the file with the given path.
 * @param {string} path - The path of the file.
 * @param {function} cb - A callback with the new workbook.
 * @returns {undefined}
 */
Workbook.fromFile = function (path, cb) {
    fs.readFile(path, function (err, data) {
        if (err) return cb(err);
        cb(null, new Workbook(data));
    });
};

/**
 * Creates a Workbook from the file with the given path synchronously.
 * @param {string} path - The path of the file.
 * @returns {Workbook} The parsed workbook.
 */
Workbook.fromFileSync = function (path) {
    var data = fs.readFileSync(path);
    return new Workbook(data);
};

/**
 * Creates a blank Workbook.
 * @param {function} cb - A callback with the new workbook.
 * @returns {undefined}
 */
Workbook.fromBlank = function (cb) {
    Workbook.fromFile(path.join(__dirname, "blank.xlsx"), cb);
};

/**
 * Creates a blank Workbook synchronously.
 * @returns {Workbook} The new workbook.
 */
Workbook.fromBlankSync = function () {
    return Workbook.fromFileSync(path.join(__dirname, "blank.xlsx"));
};

module.exports = Workbook;
