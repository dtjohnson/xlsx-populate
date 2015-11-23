"use strict";

var fs = require('fs'),
    etree = require('elementtree'),
    utils = require('./utils'),
    JSZip = require('jszip'),
    Sheet = require('./Sheet');

/**
 * Initializes a new Workbook.
 * @param {Buffer} data
 * @constructor
 */
var Workbook = function (data) {
    this._zip = new JSZip(data, { base64: false, checkCRC32: true });
    var workbookText = this._zip.file("xl/workbook.xml").asText();
    this._workbookXML = etree.parse(workbookText);

    this.sheets = [];
    var sheetNodes = this._workbookXML.findall("*/sheet");

    var removeValueNode = function (formulaParent) {
        var vNode = formulaParent.find("v");
        if (vNode) formulaParent.remove(vNode);
    };

    for (var i = 0; i < sheetNodes.length; i++) {
        var index = i + 1;
        var sheetText = this._zip.file("xl/worksheets/sheet" + index + ".xml").asText();
        var sheetXML = etree.parse(sheetText);

        // This is a blunt way to make sure formula values get updated.
        // It just clears all stored values in case the referenced cell values change.
        var formulaParents = sheetXML.findall("sheetData/row/c/f/..");
        formulaParents.forEach(removeValueNode);

        var sheet = new Sheet(this, sheetNodes[i], sheetXML);
        this.sheets.push(sheet);
    }
};

/**
 * Gets the sheet with the provided name or index (0-based).
 * @param {string|number} sheetNameOrIndex
 * @returns {Sheet}
 */
Workbook.prototype.getSheet = function (sheetNameOrIndex) {
    if (utils.isInteger(sheetNameOrIndex)) return this.sheets[sheetNameOrIndex];

    for (var i = 0; i < this.sheets.length; i++) {
        var sheet = this.sheets[i];
        if (sheet.getName() === sheetNameOrIndex) return sheet;
    }
};

/**
 * Get a named cell. (Assumes names with workbook scope pointing to single cells.)
 * @param {string} cellName
 * @returns {Cell}
 */
Workbook.prototype.getNamedCell = function (cellName) {
    var definedName = this._workbookXML.find("definedNames/definedName[@name='" + cellName + "']");
    if (!definedName) return;

    var address = definedName.text;
    var ref = utils.addressToRowAndColumn(address);
    if (!ref) return;

    return this.getSheet(ref.sheet).getCell(ref.row, ref.column);
};

/**
 * Gets the output.
 * @returns {Buffer}
 */
Workbook.prototype.output = function () {
    this._zip.file("xl/workbook.xml", this._workbookXML.write());

    for (var i = 0; i < this.sheets.length; i++) {
        var index = i + 1;
        var sheet = this.sheets[i];
        this._zip.file("xl/worksheets/sheet" + index + ".xml", sheet._sheetXML.write());
    }

    // Kill the calc chain since this will corrupt the file is formulas are removed.
    this._zip.remove("xl/calcChain.xml");

    return this._zip.generate({ type: "nodebuffer" });
};

/**
 * Writes to file with the given path.
 * @param {string} path
 * @param {function} cb
 */
Workbook.prototype.toFile = function (path, cb) {
    fs.writeFile(path, this.output(), cb);
};

/**
 * Wirtes to file with the given path synchronously.
 * @param {string} path
 */
Workbook.prototype.toFileSync = function (path) {
    fs.writeFileSync(path, this.output());
};

/**
 * Creates a Workbook from the file with the given path.
 * @param {string} path
 * @param {function} cb
 */
Workbook.fromFile = function (path, cb) {
    fs.readFile(path, function (err, data) {
        if (err) return cb(err);
        cb(null, new Workbook(data));
    });
};

/**
 * Creates a Workbook from the file with the given path synchronously.
 * @param path
 * @returns {Workbook}
 */
Workbook.fromFileSync = function (path) {
    var data = fs.readFileSync(path);
    return new Workbook(data);
};

module.exports = Workbook;
