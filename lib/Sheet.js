"use strict";

var xpath = require('./xpath');
var utils = require('./utils');
var Row = require('./Row');
var Cell = require('./Cell');

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
    this._rows = {};
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
    var getRowNode = function (sheetDataNode, rowNumber) {
        for (var childNodeIndex in sheetDataNode.childNodes) {
            if (sheetDataNode.childNodes.hasOwnProperty(childNodeIndex)) {
                var childNode = sheetDataNode.childNodes[childNodeIndex];
                if (childNode.tagName === "row") {
                    if (childNode.getAttribute("r") === String(rowNumber)) {
                        return childNode;
                    }
                }
            }
        }
        return null;
    };
    var rowNode;
    if (rowNumber in this._rows === false) {
        rowNode = getRowNode(this._sheetDataNode, rowNumber);
        if (!rowNode) {
            rowNode = this._sheetDataNode.ownerDocument.createElement("row");
            rowNode.setAttribute("r", rowNumber);
            this._sheetDataNode.appendChild(rowNode);
        }
        this._rows[rowNumber] = new Row(this, rowNode);
    }
    return this._rows[rowNumber];
};

/* eslint-disable lines-around-comment */
/**
 * Gets the cell with the given a cell.
 * @param {Cell} rowNumber - The cell.
 * @returns {Cell} The cell.
 *//**
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
    if (arguments.length === 0) {
        throw new Error("Row and column must be provided");
    }
    if (arguments[0] instanceof Cell) {
        return arguments[0];
    }
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
    if (utils.isInteger(rowNumber) === false || rowNumber <= 0) {
        throw new Error("Row must be a positive integer");
    }
    if (utils.isInteger(columnNumber) === false || columnNumber <= 0) {
        throw new Error("Column must be a positive integer");
    }
    return this.getRow(rowNumber).getCell(columnNumber);
};

module.exports = Sheet;
