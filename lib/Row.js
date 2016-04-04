"use strict";

var xpath = require('./xpath');
var utils = require('./utils');
var Cell = require('./Cell');

/**
 * Initializes a new Row.
 * @param {Sheet} sheet - The parent sheet.
 * @param {Element} rowNode - The row's node.
 * @constructor
 */
var Row = function (sheet, rowNode) {
    this._sheet = sheet;
    this._rowNode = rowNode;
};

/**
 * Gets the parent sheet.
 * @returns {Sheet} The parent sheet.
 */
Row.prototype.getSheet = function () {
    return this._sheet;
};

/**
 * Gets the row number of the row.
 * @returns {number} The row number.
 */
Row.prototype.getRowNumber = function () {
    return parseInt(this._rowNode.getAttribute("r"));
};

/**
 * Gets the cell in the row with the provided column number.
 * @param {number} columnNumber - The column number.
 * @returns {Cell} The cell with the provided column number.
 */
Row.prototype.getCell = function (columnNumber) {
    var address = utils.rowAndColumnToAddress(this.getRowNumber(), columnNumber);
    var cellNode = xpath("(sml:c[@r='" + address + "'])[1]", this._rowNode)[0];
    if (!cellNode) {
        cellNode = this._rowNode.ownerDocument.createElement("c");
        cellNode.setAttribute("r", address);
        this._rowNode.appendChild(cellNode);
    }

    return new Cell(this, cellNode);
};

module.exports = Row;
