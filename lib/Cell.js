"use strict";

/* eslint no-console: ["error", { allow: ["log", warn", "error"] }] */

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
 * Get node information.
 * @returns {string} The cell information.
 */
Cell.prototype.toString = function () {
    var buffer = [];
    buffer.push('address: ' + this.getFullAddress());
    return utils.getNodeInfo(this._cellNode, buffer);
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

    var isNode, tNode, vNode, textNode;
    if (typeof value === "string") {
        this._cellNode.setAttribute("t", "inlineStr");
        isNode = this._cellNode.ownerDocument.createElement("is");
        this._cellNode.appendChild(isNode);
        tNode = this._cellNode.ownerDocument.createElement("t");
        isNode.appendChild(tNode);
        textNode = this._cellNode.ownerDocument.createTextNode(value);
        tNode.appendChild(textNode);
    } else if (typeof value === "boolean") {
        this._cellNode.setAttribute("t", "b");
        vNode = this._cellNode.ownerDocument.createElement("v");
        this._cellNode.appendChild(vNode);
        textNode = this._cellNode.ownerDocument.createTextNode(value ? 1 : 0);
        vNode.appendChild(textNode);
    } else if (typeof value === "number") {
        vNode = this._cellNode.ownerDocument.createElement("v");
        this._cellNode.appendChild(vNode);
        textNode = this._cellNode.ownerDocument.createTextNode(value);
        vNode.appendChild(textNode);
    } else if (value instanceof Date) {
        vNode = this._cellNode.ownerDocument.createElement("v");
        this._cellNode.appendChild(vNode);
        textNode = this._cellNode.ownerDocument.createTextNode(utils.dateToExcelNumber(value));
        vNode.appendChild(textNode);
    }

    return this;
};

/**
 * Returns a cell with a relative position to the offsets provided.
 * @param {number} rowOffset - Offset from this.getRowNumber().
 * @param {number} columnOffset - Offset from this.getColumnNumber().
 * @returns {Cell} The relative cell.
 */
Cell.prototype.getRelativeCell = function (rowOffset, columnOffset) {
    if (!utils.isInteger(rowOffset)) {
        console.error(this.toString());
        throw new Error(rowOffset, 'Expected row offset to be an integer');
    }
    if (!utils.isInteger(columnOffset)) {
        console.error(this.toString());
        throw new Error(columnOffset, 'Expected column offset to be an integer');
    }
    var absoluteRow = rowOffset + this.getRowNumber();
    if (absoluteRow < 0) {
        console.error(this.toString());
        throw new Error(absoluteRow, 'Expected relative row to be a non-negative integer');
    }
    var absoluteColumn = columnOffset + this.getColumnNumber();
    if (absoluteColumn < 0) {
        console.error(this.toString());
        throw new Error(absoluteColumn, 'Expected relative column to be a non-negative integer');
    }
    return this.getSheet().getCell(absoluteRow, absoluteColumn);
};

/**
 * Sets the formula for a cell (with optional pre-calculated value).
 * @param {string} formula - The formula to set.
 * @param {*} [calculatedValue] - The pre-calculated value.
 * @param {number} [sharedIndex] - Unique non-negative integer value to represent the formula.
 * @param {string} [sharedRef] - Range of cells referencing this formala, for example: "A1:A3".
 * @returns {Cell} The cell.
 */
Cell.prototype.setFormula = function (formula, calculatedValue, sharedIndex, sharedRef) {
    this.setValue(calculatedValue);

    var fNode = this._cellNode.ownerDocument.createElement('f');
    this._cellNode.appendChild(fNode);

    if (typeof formula === 'string') {
        if (formula.length > 0) {
            var textNode = this._cellNode.ownerDocument.createTextNode(formula);
            fNode.appendChild(textNode);
        }
    }

    if (utils.isInteger(sharedIndex)) {
        if (sharedIndex >= 0) {
            // TODO: Ensure that sharedIndex is unique
            fNode.setAttribute('t', 'shared');
            fNode.setAttribute('si', String(sharedIndex));
        }
    }

    if (typeof sharedRef === 'string') {
        fNode.setAttribute('ref', sharedRef);
    }

    return this;
};

/**
 * Determine whether the cell is a shared formula.
 * @param {boolean} [isSource] - If true, also check for formula definition.
 * @returns {boolean} The is shared formula boolean.
 */
Cell.prototype.isSharedFormula = function (isSource) {
    isSource = isSource || false;

    /* XLSX structure of shared formulas:
    <sheetData>
        <row ...>
            <c ...>
                <f ref="F2:F519" si="0" t="shared">C2/B2</f>
                <f si="0" t="shared" />
                ...
            </c>
        </row>
    </sheetData>
    */

    var fNode = this._cellNode.getElementsByTagName('f')[0];
    if (!fNode) {
        console.log(this.toString(), '<f> node not found');
        console.log(utils.getNodeInfo(this._cellNode));
        return false;
    }
    if (isSource) {
        var fNodeText = utils.getNodeText(fNode);
        if (!fNodeText || !fNodeText.length) {
            console.log(this.toString(), '<f> node formula is empty');
            return false;
        }
        var fNodeRef = fNode.getAttribute('ref');
        if (!fNodeRef || !fNodeRef.length) {
            console.log(utils.getNodeInfo(fNode));
            console.log(this.toString(), '<f> node ref (address range) is empty');
            return false;
        }
    }
    var fNodeType = fNode.getAttribute('t');
    if (fNodeType !== 'shared') {
        console.log(utils.getNodeInfo(fNode));
        console.log(this.toString(), '<f> node t (type) not shared');
        return false;
    }
    var fNodeSharedIndex = fNode.getAttribute('si');
    if (!fNodeSharedIndex || !fNodeSharedIndex.length) {
        console.log(utils.getNodeInfo(fNode));
        console.log(this.toString(), '<f> node si (shared index) is empty');
        return false;
    }
    return true;
};

/**
 * If this cell is the source of a shared formula,
 * then assign all cells from this cell to lastSharedCell its shared index.
 * Note that lastSharedCell must share the same row or column, such that
 *   this.getRowNumber() <= lastSharedCell.getRowNumber()
 *       AND,
 *   this.getColumnNumber() <= lastSharedCell.getColumnNumber()
 * @param {*} lastSharedCell - String address or cell to share formula up until.
 * @returns {Cell} The shared formula source cell.
 */
Cell.prototype.shareFormulaUntil = function (lastSharedCell) {
    if (this.isSharedFormula(true) === false) {
        throw new Error('Expected cell to be a shared formula source');
    }
    var fNode = this._cellNode.getElementsByTagName('f')[0];
    var sharedIndex = parseInt(fNode.getAttribute('si'));
    if (!utils.isInteger(sharedIndex) || sharedIndex < 0) {
        console.error(utils.getNodeInfo(fNode));
        console.error(this.toString());
        throw new Error('Expected shared index to be a non-negative integer');
    }
    if (typeof lastSharedCell === 'string') {
        lastSharedCell = this.getSheet().getCell(lastSharedCell);
    }
    if (lastSharedCell instanceof Cell === false) {
        console.error('lastSharedCell: ', lastSharedCell);
        console.error(this.toString());
        throw new Error('Expected lastSharedCell to be a cell');
    }
    var cell;
    var rowNumber = this.getRowNumber();
    var columnNumber = this.getColumnNumber();
    var lastSharedCellRowNumber = lastSharedCell.getRowNumber();
    var lastSharedCellColumnNumber = lastSharedCell.getColumnNumber();
    if (rowNumber === lastSharedCellRowNumber) {
        for (var c = 1 + columnNumber; c <= lastSharedCellColumnNumber; c++) {
            this
                .getSheet()
                .getCell(rowNumber, c)
                .setFormula(undefined, undefined, sharedIndex)
                ;
        }
    } else if (columnNumber === lastSharedCellColumnNumber) {
        for (var r = 1 + rowNumber; r <= lastSharedCellRowNumber; r++) {
            this
                .getSheet()
                .getCell(r, columnNumber)
                .setFormula(undefined, undefined, sharedIndex)
                ;
        }
    } else {
        console.error('lastSharedCell: ', lastSharedCell.toString());
        console.error(this.toString());
        throw new Error('Expected last shared forumla cell to align either row-wise or column-wise with shared formula source');
    }
    var sharedRef = this.getAddress() + ':' + lastSharedCell.getAddress();
    fNode.setAttribute('ref', sharedRef);
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
