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
    this._sheet._cacheCells = this._sheet._cacheCells || {};
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

    var getNodeColumnNumber = function (node) {
        return utils.addressToRowAndColumn(node.getAttribute("r")).column;
    };

    // Find the matching child node or the next node. Don't use xpath as it's too slow.
    // Nodes must be in order!
    var nextNode;
    var searchResult = {
        found: false,
        index: 0
    };

    // Memoization
    if (!searchResult.found) {
        if (columnNumber in this._sheet._cacheCells) {
            var cachedIndex = this._sheet._cacheCells[columnNumber];
            var cachedNode = this._rowNode.childNodes[cachedIndex];
            if (cachedNode) {
                if (getNodeColumnNumber(cachedNode) === columnNumber) {
                    searchResult = { found: true, index: cachedIndex };
                }
            }
        }
    }

    // Binary search
    if (!searchResult.found) {
        if (this._rowNode.hasChildNodes()) {
            searchResult = utils.binarySearch(columnNumber, this._rowNode.childNodes, getNodeColumnNumber);
        }
    }

    nextNode = this._rowNode.childNodes[searchResult.index];

    var cellNode;
    if (searchResult.found) {
        cellNode = nextNode;
    } else {
        // No existing node so create a new one.
        cellNode = this._rowNode.ownerDocument.createElement("c");
        cellNode.setAttribute("r", address);

        // Insert or append the new node.
        if (nextNode) {
            this._rowNode.insertBefore(cellNode, nextNode);
        } else {
            this._rowNode.appendChild(cellNode);
        }
    }

    this._sheet._cacheCells[columnNumber] = searchResult.index;

    return new Cell(this, cellNode);
};

module.exports = Row;
