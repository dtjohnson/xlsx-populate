"use strict";

const utils = require('./utils');
const debug = require('./debug');

/**
 * A workbook cell.
 */
class Cell {
    /**
     * Initializes a new Cell.
     * @param {Row} row - The parent row.
     * @param {Element} cellNode - The cell node.
     * @constructor
     * @private
     */
    constructor(row, cellNode) {
        this._row = row;
        this._cellNode = cellNode;
    }

    /**
     * Gets the address of the cell (e.g. "A5").
     * @returns {string} The cell address.
     */
    address() {
        if (arguments.length > 0) throw new Error('Cell.address cannot be set.');
        return this._cellNode.getAttribute("r");
    }

    /**
     * Gets the column name of the cell.
     * @returns {number} The column name.
     */
    columnName() {
        if (arguments.length > 0) throw new Error('Cell.columnName cannot be set.');
        return utils.columnNumberToName(this.columnNumber());
    }

    /**
     * Gets the column number of the cell.
     * @returns {number} The column number.
     */
    columnNumber() {
        if (arguments.length > 0) throw new Error('Cell.columnNumber cannot be set.');
        return utils.addressToRowAndColumn(this.address()).column;
    }

    /**
     * Gets the parent row of the cell.
     * @returns {Row} The parent row.
     */
    row() {
        if (arguments.length > 0) throw new Error('Cell.row cannot be set.');
        return this._row;
    }

    /**
     * Gets the row number of the cell.
     * @returns {number} The row number.
     */
    rowNumber() {
        if (arguments.length > 0) throw new Error('Cell.rowNumber cannot be set.');
        return utils.addressToRowAndColumn(this.address()).row;
    }

    /**
     * Gets the parent sheet.
     * @returns {Sheet} The parent sheet.
     */
    sheet() {
        if (arguments.length > 0) throw new Error('Cell.sheet cannot be set.');
        return this.row().sheet();
    }

    /**
     * Gets the parent workbook.
     * @returns {Workbook} The parent workbook.
     */
    workbook() {
        if (arguments.length > 0) throw new Error('Cell.workbook cannot be set.');
        return this.row().workbook();
    }












    /**
     * Gets the full address of the cell including sheet (e.g. "Sheet1!A5").
     * @returns {string} The full address.
     */
    fullAddress() {
        return utils.addressToFullAddress(this.sheet().name(), this.address());
    }

    /**
     * Sets the value of the cell.
     * @param {*} value - The value to set.
     * @returns {Cell} The cell.
     */
    value(value) {
        this.clear();

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
    }

    /**
     * Returns a cell with a relative position to the offsets provided.
     * @param {number} rowOffset - Offset from this.rowNumber().
     * @param {number} columnOffset - Offset from this.columnNumber().
     * @returns {Cell} The relative cell.
     */
    relativeCell(rowOffset, columnOffset) {
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
        var absoluteRow = rowOffset + this.rowNumber();
        if (absoluteRow < 0) {
            throw new Error(
                this,
                'Expected relative row to be a non-negative integer',
                absoluteRow
            );
        }
        var absoluteColumn = columnOffset + this.columnNumber();
        if (absoluteColumn < 0) {
            throw new Error(
                this,
                'Expected relative column to be a non-negative integer',
                absoluteColumn
            );
        }
        return this.sheet().cell(absoluteRow, absoluteColumn);
    }

    /**
     * Sets the formula for a cell (with optional pre-calculated value).
     * @param {string} formula - The formula to set.
     * @param {*} [calculatedValue] - The pre-calculated value.
     * @param {number} [sharedIndex] - Unique non-negative integer value to represent the formula.
     * @param {string} [sharedRef] - Range of cells referencing this formala, for example: "A1:A3".
     * @returns {Cell} The cell.
     */
    formula(formula, calculatedValue, sharedIndex, sharedRef) {
        this.value(calculatedValue);

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
    }

    /**
     * Determine whether the cell is a shared formula.
     * @param {boolean} [isSource] - If true, also check for formula definition.
     * @returns {boolean} The is shared formula boolean.
     * @private
     */
    _isSharedFormula(isSource) {
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
    }

    /**
     * If this cell is the source of a shared formula,
     * then assign all cells from this cell to lastSharedCell its shared index.
     * Note that lastSharedCell must share the same row or column, such that
     *   this.rowNumber() <= lastSharedCell.rowNumber()
     *       AND,
     *   this.columnNumber() <= lastSharedCell.columnNumber()
     * @param {*} lastSharedCell - String address or cell to share formula up until.
     * @returns {Cell} The shared formula source cell.
     */
    shareFormulaUntil(lastSharedCell) {
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
            lastSharedCell = this.sheet().cell(lastSharedCell);
        }
        if (lastSharedCell instanceof Cell === false) {
            throw new Error(
                this,
                'Expected lastSharedCell to be a cell',
                lastSharedCell
            );
        }
        var cell;
        var rowNumber = this.rowNumber();
        var columnNumber = this.columnNumber();
        var lastSharedCellRowNumber = lastSharedCell.rowNumber();
        var lastSharedCellColumnNumber = lastSharedCell.columnNumber();
        if (rowNumber === lastSharedCellRowNumber) {
            for (var c = 1 + columnNumber; c <= lastSharedCellColumnNumber; c++) {
                this
                    .sheet()
                    .cell(rowNumber, c)
                    .formula(undefined, undefined, sharedIndex)
                    ;
            }
        } else if (columnNumber === lastSharedCellColumnNumber) {
            for (var r = 1 + rowNumber; r <= lastSharedCellRowNumber; r++) {
                this
                    .sheet()
                    .cell(r, columnNumber)
                    .formula(undefined, undefined, sharedIndex)
                    ;
            }
        } else {
            throw new Error(
                this,
                'Expected last shared forumla cell to align either row-wise or column-wise with shared formula source',
                lastSharedCell
            );
        }
        var sharedRef = this.address() + ':' + lastSharedCell.address();
        fNode.setAttribute('ref', sharedRef);
        return this;
    }

    /**
     * Clears the contents from the cell.
     * @returns {undefined}
     * @private
     */
    clear() {
        while (this._cellNode.firstChild) {
            this._cellNode.removeChild(this._cellNode.firstChild);
        }

        this._cellNode.removeAttribute("t");
    }

    /**
     * Get node information.
     * @returns {string} The cell information.
     */
    toString() {
        return utils.getNodeInfo(this._cellNode, {
            address: this.fullAddress()
        });
    }
}

module.exports = Cell;
