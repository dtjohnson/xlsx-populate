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

    // Find the matching child node or the next node. Don't use xpath as it's too slow.
    // Nodes must be in order!
    var cellNode, nextNode;
    for (var i = 0; i < this._rowNode.childNodes.length; i++) {
        var childNode = this._rowNode.childNodes[i];
        var r = childNode.getAttribute("r");
        if (r === address) {
            cellNode = childNode;
            break;
        } else if (utils.addressToRowAndColumn(r).column > columnNumber) {
            nextNode = childNode;
            break;
        }
    }

    // No existing node so create a new one.
    if (!cellNode) {
        cellNode = this._rowNode.ownerDocument.createElement("c");
        cellNode.setAttribute("r", address);

        // Insert or append the new node.
        if (nextNode) {
            this._rowNode.insertBefore(cellNode, nextNode);
        } else {
            this._rowNode.appendChild(cellNode);
        }
    }

    return new Cell(this, cellNode);
};

module.exports = Row;
