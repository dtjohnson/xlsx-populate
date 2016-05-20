"use strict";

var xpath = require('./xpath');
var utils = require('./utils');
var Row = require('./Row');

/**
 * Initializes a new Sheet.
 * @param {Workbook} workbook - The parent workbook.
 * @param {Element} sheetNode - The node defining the sheet metadat in the workbook.xml.
 * @param {Document} sheetXML - The XML defining the sheet data in worksheets/sheetN.xml.
 * @constructor
 */
var Sheet = function (workbook, sheetNode, sheetXML) {
    this._workbook = workbook;
    this._sheetNode = sheetNode;
    this._sheetXML = sheetXML;
    this._sheetDataNode = xpath('sml:sheetData', sheetXML)[0];
    this._cacheRows = {};
};

/**
 * Gets the parent workbook.
 * @returns {Workbook} The parent workbook.
 */
Sheet.prototype.getWorkbook = function () {
    return this._workbook;
};

/**
 * Gets the name of the sheet.
 * @returns {string} The name of the sheet.
 */
Sheet.prototype.getName = function () {
    return this._sheetNode.getAttribute("name");
};

/**
 * Set the name of the sheet.
 * @param {string} name - The new name of the sheet.
 * @returns {undefined}
 */
Sheet.prototype.setName = function (name) {
    this._sheetNode.setAttribute("name", name);
};

/**
 * Gets the row with the given number.
 * @param {number} rowNumber - The row number.
 * @returns {Row} The row with the given number.
 */
Sheet.prototype.getRow = function (rowNumber) {
    var getNodeRowNumber = function (node) {
        return parseInt(node.getAttribute("r"));
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
        if (rowNumber in this._cacheRows) {
            var cachedIndex = this._cacheRows[rowNumber];
            var cachedNode = this._sheetDataNode.childNodes[cachedIndex];
            if (cachedNode) {
                if (getNodeRowNumber(cachedNode) === rowNumber) {
                    searchResult = { found: true, index: cachedIndex };
                }
            }
        }
    }

    // Binary search
    if (!searchResult.found) {
        if (this._sheetDataNode.hasChildNodes()) {
            searchResult = utils.binarySearch(rowNumber, this._sheetDataNode.childNodes, getNodeRowNumber);
        }
    }

    nextNode = this._sheetDataNode.childNodes[searchResult.index];

    var rowNode;
    if (searchResult.found) {
        rowNode = nextNode;
    } else {
        // No existing node so create a new one.
        rowNode = this._sheetDataNode.ownerDocument.createElement("row");
        rowNode.setAttribute("r", rowNumber);

        // Insert or append the new node.
        if (nextNode) {
            this._sheetDataNode.insertBefore(rowNode, nextNode);
        } else {
            this._sheetDataNode.appendChild(rowNode);
        }
    }

    this._cacheRows[rowNumber] = searchResult.index;

    return new Row(this, rowNode);
};

/* eslint-disable lines-around-comment */
/**
 * Gets the cell with the given address.
 * @param {string} address - The address of the cell.
 * @returns {Cell} The cell.
*//**
 * Gets the cell with the given row and column numbers.
 * @param {number} rowNumber - The row number of the cell.
 * @param {number} columnNumber - The column number of the cell.
 * @returns {Cell} The cell.
 */
/* eslint-enable lines-around-comment */
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
