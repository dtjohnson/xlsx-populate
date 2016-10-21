"use strict";


var utils = require('./utils');
var debug = require('debug')('Cell');

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
    return utils.getNodeInfo(this._cellNode, {
        address: this.getFullAddress()
    });
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
        throw new Error(
            this,
            'Expected row offset to be an integer',
            rowOffset
        );
    }
    if (!utils.isInteger(columnOffset)) {
        throw new Error(
            this,
            'Expected column offset to be an integer',
            columnOffset
        );
    }
    var absoluteRow = rowOffset + this.getRowNumber();
    if (absoluteRow < 0) {
        throw new Error(
            this,
            'Expected relative row to be a non-negative integer',
            absoluteRow
        );
    }
    var absoluteColumn = columnOffset + this.getColumnNumber();
    if (absoluteColumn < 0) {
        throw new Error(
            this,
            'Expected relative column to be a non-negative integer',
            absoluteColumn
        );
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
 * @private
 */
Cell.prototype._isSharedFormula = function (isSource) {
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
        debug('Cell %s', this);
        debug('Node <f> (formula) not found');
        debug('Node %s', utils.getNodeInfo(this._cellNode));
        return false;
    }
    if (isSource) {
        var fNodeText = utils.getNodeText(fNode);
        if (!fNodeText || !fNodeText.length) {
            debug('Cell %s', this);
            debug('Node <f> (formula) is empty');
            return false;
        }
        var fNodeRef = fNode.getAttribute('ref');
        if (!fNodeRef || !fNodeRef.length) {
            debug('Cell %s', this);
            debug('Node <f> (formula) attribute ref (address range) is empty');
            debug('Node <f> %s', utils.getNodeInfo(fNode));
            return false;
        }
    }
    var fNodeType = fNode.getAttribute('t');
    if (fNodeType !== 'shared') {
        debug('Cell %s', this);
        debug('Node <f> (formula) attribute t (type) not shared');
        debug('Node <f> %s', utils.getNodeInfo(fNode));
        return false;
    }
    var fNodeSharedIndex = fNode.getAttribute('si');
    if (!fNodeSharedIndex || !fNodeSharedIndex.length) {
        debug('Cell %s', this);
        debug('Node <f> (formula) attribute si (shared index) is empty');
        debug('Node <f> %s', utils.getNodeInfo(fNode));
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
    if (this._isSharedFormula(true) === false) {
        throw new Error('Expected cell to be a shared formula source');
    }
    var fNode = this._cellNode.getElementsByTagName('f')[0];
    var sharedIndex = parseInt(fNode.getAttribute('si'));
    if (!utils.isInteger(sharedIndex) || sharedIndex < 0) {
        throw new Error(
            this,
            'Expected shared index to be a non-negative integer',
            utils.getNodeInfo(fNode)
        );
    }
    if (typeof lastSharedCell === 'string') {
        lastSharedCell = this.getSheet().getCell(lastSharedCell);
    }
    if (lastSharedCell instanceof Cell === false) {
        throw new Error(
            this,
            'Expected lastSharedCell to be a cell',
            lastSharedCell
        );
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
        throw new Error(
            this,
            'Expected last shared forumla cell to align either row-wise or column-wise with shared formula source',
            lastSharedCell
        );
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
