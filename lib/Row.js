"use strict";

var etree = require('elementtree'),
    subelement = etree.SubElement,
    utils = require('./utils'),
    Cell = require('./Cell');

/**
 * Initializes a new Row.
 * @param {Sheet} sheet
 * @param {number} rowNumber
 * @param {etree.SubElement} rowNode
 * @constructor
 */
var Row = function (sheet, rowNumber, rowNode) {
    this._sheet = sheet;
    this._rowNumber = rowNumber;
    this._rowNode = rowNode;
};

/**
 * Gets the parent sheet.
 * @returns {Sheet}
 */
Row.prototype.getSheet = function () {
    return this._sheet;
};

/**
 * Gets the row number of the row.
 * @returns {number}
 */
Row.prototype.getRowNumber = function () {
    return this._rowNumber;
};

/**
 * Gets the cell in the row with the provided column number.
 * @param {number} columnNumber - The column number.
 * @returns {Cell} The cell with the provided column number.
 */
Row.prototype.getCell = function (columnNumber) {
    var address = utils.rowAndColumnToAddress(this._rowNumber, columnNumber);
    var cellNode = this._rowNode.find("c[@r='" + address + "']");
    if (!cellNode) {
        cellNode = subelement(this._rowNode, "c");
        cellNode.attrib.r = address;
    }

    return new Cell(this.getSheet(), this, columnNumber, cellNode);
};

module.exports = Row;
