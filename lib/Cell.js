"use strict";

var utils = require('./utils');

/**
 * Initializes a new Cell.
 * @param {Row} row - The parent row.
 * @param {Element} cellNode - The cell node.
 * @constructor
 */
var Cell = function (row, cellNode) {
    this._row = row;
    this._cellNode = cellNode;
};

/**
 * Gets the parent row of the cell.
 * @returns {Row} The parent row.
 */
Cell.prototype.getRow = function () {
    return this._row;
};

/**
 * Gets the parent sheet.
 * @returns {Sheet} The parent sheet.
 */
Cell.prototype.getSheet = function () {
    return this.getRow().getSheet();
};

/**
 * Gets the address of the cell (e.g. "A5").
 * @returns {string} The cell address.
 */
Cell.prototype.getAddress = function () {
    return this._cellNode.getAttribute("r");
};

/**
 * Gets the row number of the cell.
 * @returns {number} The row number.
 */
Cell.prototype.getRowNumber = function () {
    return utils.addressToRowAndColumn(this.getAddress()).row;
};

/**
 * Gets the column number of the cell.
 * @returns {number} The column number.
 */
Cell.prototype.getColumnNumber = function () {
    return utils.addressToRowAndColumn(this.getAddress()).column;
};

/**
 * Gets the column name of the cell.
 * @returns {number} The column name.
 */
Cell.prototype.getColumnName = function () {
    return utils.columnNumberToName(this.getColumnNumber());
};


/**
 * Gets the full address of the cell including sheet (e.g. "Sheet1!A5").
 * @returns {string} The full address.
 */
Cell.prototype.getFullAddress = function () {
    return utils.addressToFullAddress(this.getSheet().getName(), this.getAddress());
};

/**
 * Sets the value of the cell.
 * @param {*} value - The value to set.
 * @returns {Cell} The cell.
 */
Cell.prototype.setValue = function (value) {
    this._clearContents();

    var t, text, innerTag;
    if (typeof value === "string") {
        t = "inlineStr";
        innerTag = "is";
        text = value;
    } else if (typeof value === "boolean") {
        t = "b";
        innerTag = "v";
        text = value ? 1 : 0;
    } else if (typeof value === "number") {
        innerTag = "v";
        text = value;
    } else if (value instanceof Date) {
        // TODO
        throw new Error("Not yet supported.");
    }

    if (t) this._cellNode.setAttribute("t", t);
    if (innerTag) {
        var innerNode = this._cellNode.ownerDocument.createElement(innerTag);
        this._cellNode.appendChild(innerNode);
        var textNode = this._cellNode.ownerDocument.createTextNode(text);
        innerNode.appendChild(textNode);
    }

    return this;
};

/**
 * Sets the formula for a cell (with optional pre-calculated value).
 * @param {string} formula - The formula to set.
 * @param {*} [calculatedValue] - The pre-calculated value.
 * @returns {Cell} The cell.
 */
Cell.prototype.setFormula = function (formula, calculatedValue) {
    this.setValue(calculatedValue);

    var fNode = this._cellNode.ownerDocument.createElement("f");
    this._cellNode.appendChild(fNode);
    var textNode = this._cellNode.ownerDocument.createTextNode(formula);
    fNode.appendChild(textNode);

    return this;
};

/**
 * Clears the contents from the cell.
 * @returns {undefined}
 * @private
 */
Cell.prototype._clearContents = function () {
    while (this._cellNode.firstChild) {
        this._cellNode.removeChild(this._cellNode.firstChild);
    }

    this._cellNode.removeAttribute("t");
};

module.exports = Cell;
