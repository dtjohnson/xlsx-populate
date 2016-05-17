"use strict";

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
    this._cells = {};
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
    var getCellNode = function (rowNode, address) {
        for (var childNodeIndex in rowNode.childNodes) {
            if (rowNode.childNodes.hasOwnProperty(childNodeIndex)) {
                var childNode = rowNode.childNodes[childNodeIndex];
                if (childNode.tagName === "c") {
                    if (childNode.getAttribute("r") === address) {
                        return childNode;
                    }
                }
            }
        }
        return null;
    };
    var cellNode;
    if (address in this._cells === false) {
        cellNode = getCellNode(this._rowNode, address);
        if (!cellNode) {
            cellNode = this._rowNode.ownerDocument.createElement("c");
            cellNode.setAttribute("r", address);
            this._rowNode.appendChild(cellNode);
        }
        this._cells[address] = new Cell(this, cellNode);
    }
    return this._cells[address];
};

module.exports = Row;
