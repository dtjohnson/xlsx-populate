"use strict";

var utils = require('./utils'),
    etree = require('elementtree'),
    subelement = etree.SubElement,
    Row = require('./Row');

/**
 * Initializes a new Sheet.
 * @param {Workbook} workbook
 * @param {string} name
 * @param {etree.Element} sheetNode - The node defining the sheet in the workbook.xml.
 * @param {etree.Element} sheetXML
 * @constructor
 */
var Sheet = function (workbook, sheetNode, sheetXML) {
    this._workbook = workbook;
    this._sheetNode = sheetNode;
    this._sheetXML = sheetXML;
    this._sheetDataNode = this._sheetXML.find("sheetData");
};

/**
 * Gets the parent workbook.
 * @returns {Workbook}
 */
Sheet.prototype.getWorkbook = function () {
    return this._workbook;
};

/**
 * Gets the name of the sheet.
 * @returns {string}
 */
Sheet.prototype.getName = function () {
    return this._sheetNode.attrib.name;
};

Sheet.prototype.setName = function (name) {
    this._sheetNode.attrib.name = name;
};

Sheet.prototype.getRow = function (rowNumber) {
    var rowNode = this._sheetDataNode.find("row[@r='" + rowNumber + "']");
    if (!rowNode) {
        rowNode = subelement(this._sheetDataNode, "row");
        rowNode.attrib.r = String(rowNumber); // Convert to string so we can find the row later
    }

    return new Row(this, rowNumber, rowNode);
};

/**
 * Gets the cell with either the provided row and column or address.
 * @returns {Cell}
 */
Sheet.prototype.getCell = function () {
    var rowNumber, columnNumber;
    if (arguments.length === 1) {
        var address = arguments[0];
        var ref = utils.addressToRowAndColumn(address);
        rowNumber = ref.row;
        columnNumber = ref.column;
    } else {
        rowNumber = arguments[0];
        columnNumber = arguments[1];
    }

    return this.getRow(rowNumber).getCell(columnNumber);
};

module.exports = Sheet;
