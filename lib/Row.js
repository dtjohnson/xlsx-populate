"use strict";

const _ = require("lodash");
const Cell = require("./Cell");
const regexify = require("./regexify");
const debug = require("./debug")('Row');
const ArgHandler = require("./ArgHandler");
const addressConverter = require('./addressConverter');

/**
 * A row.
 */
class Row {
    // /**
    //  * Creates a new instance of Row.
    //  * @param {Sheet} sheet - The parent sheet.
    //  * @param {{}} node - The row node.
    //  */
    constructor(sheet, node) {
        debug("constructor(...)");
        this._sheet = sheet;
        this._init(node);
    }

    /**
     * Get the address of the row.
     * @param {{}} [opts] - Options
     * @param {boolean} [opts.includeSheetName] - Include the sheet name in the address.
     * @param {boolean} [opts.anchored] - Anchor the address.
     * @returns {string} The address
     */
    address(opts) {
        debug("address(%o)", arguments);
        return addressConverter.toAddress({
            type: 'row',
            rowNumber: this.rowNumber(),
            sheetName: opts && opts.includeSheetName && this.sheet().name(),
            rowAnchored: opts && opts.anchored
        });
    }

    /**
     * Get a cell in the row.
     * @param {string|number} columnNameOrNumber - The name or number of the column.
     * @returns {Cell} The cell.
     */
    cell(columnNameOrNumber) {
        debug("cell(%o)", arguments);
        let columnNumber = columnNameOrNumber;
        if (typeof columnNameOrNumber === 'string') {
            columnNumber = addressConverter.columnNameToNumber(columnNameOrNumber);
        }

        if (this._cells[columnNumber]) return this._cells[columnNumber];

        const address = addressConverter.toAddress({
            type: 'cell',
            rowNumber: this.rowNumber(),
            columnNumber
        });

        const cellNode = { name: 'c', attributes: { r: address }, children: [] };

        // Copy existing row/column styles to the new cell.
        const columnStyleId = this.sheet().existingColumnStyleId(columnNumber);
        const rowStyleId = this._node.attributes.s;
        if (!_.isNil(columnStyleId)) cellNode.attributes.s = columnStyleId;
        else if (!_.isNil(rowStyleId)) cellNode.attributes.s = rowStyleId;

        const cell = new Cell(this, cellNode);
        this._cells[columnNumber] = cell;
        return cell;
    }

    /**
     * Gets the row height.
     * @returns {undefined|number} The height (or undefined).
     *//**
     * Sets the row height.
     * @param {number} height - The height of the row.
     * @returns {Row} The row.
     */
    height() {
        debug('height(%o)', arguments);
        return new ArgHandler('Row.height')
            .case(() => {
                return this._node.attributes.customHeight ? this._node.attributes.ht : undefined;
            })
            .case('number', height => {
                this._node.attributes.ht = height;
                this._node.attributes.customHeight = 1;
                return this;
            })
            .case('nil', () => {
                delete this._node.attributes.ht;
                delete this._node.attributes.customHeight;
                return this;
            })
            .handle(arguments);
    }

    /**
     * Gets a value indicating whether the row is hidden.
     * @returns {boolean} A flag indicating whether the row is hidden.
     *//**
     * Sets whether the row is hidden.
     * @param {boolean} hidden - A flag indicating whether to hide the row.
     * @returns {Row} The row.
     */
    hidden() {
        debug('hidden(%o)', arguments);
        return new ArgHandler("Row.hidden")
            .case(() => {
                return this._node.attributes.hidden === 1;
            })
            .case('boolean', hidden => {
                if (hidden) this._node.attributes.hidden = 1;
                else delete this._node.attributes.hidden;
                return this;
            })
            .handle(arguments);
    }

    /**
     * Gets the row number.
     * @returns {number} The row number.
     */
    rowNumber() {
        debug("rowNumber(%o)", arguments);
        return this._node.attributes.r;
    }

    /**
     * Gets the parent sheet of the row.
     * @returns {Sheet} The parent sheet.
     */
    sheet() {
        debug("sheet(%o)", arguments);
        return this._sheet;
    }

    /**
     * Get the parent workbook.
     * @returns {XlsxPopulate} The parent workbook.
     */
    workbook() {
        debug("workbook(%o)", arguments);
        return this.sheet().workbook();
    }

    /**
     * Clear cells that are using a given shared formula ID.
     * @param {number} sharedFormulaId - The shared formula ID.
     * @returns {undefined}
     * @ignore
     */
    clearCellsUsingSharedFormula(sharedFormulaId) {
        debug("clearCellsUsingSharedFormula(%o)", arguments);
        this._cells.forEach(cell => {
            if (!cell) return;
            if (cell.sharesFormula(sharedFormulaId)) cell.clear();
        });
    }

    /**
     * Find a pattern in the row and optionally replace it.
     * @param {string|RegExp} pattern - The search pattern.
     * @param {string} [replacement] - The replacement text.
     * @returns {Array.<Cell>} The matched cells.
     * @ignore
     */
    find(pattern, replacement) {
        debug("find(%o)", arguments);
        pattern = regexify(pattern);

        const matches = [];
        this._cells.forEach(cell => {
            if (!cell) return;
            if (cell.find(pattern, replacement)) matches.push(cell);
        });

        return matches;
    }

    /**
     * Returns the nax used column number.
     * @returns {number} The max used column number.
     * @ignore
     */
    minUsedColumnNumber() {
        debug("minUsedColumnNumber(%o)", arguments);
        return _.findIndex(this._cells);
    }

    /**
     * Returns the nax used column number.
     * @returns {number} The max used column number.
     * @ignore
     */
    maxUsedColumnNumber() {
        debug("maxUsedColumnNumber(%o)", arguments);
        return this._cells.length - 1;
    }

    /**
     * Convert the row to an object.
     * @returns {{}} The object form.
     * @ignore
     */
    toObject() {
        debug("toObject(%o)", arguments);

        // Cells must be in order.
        this._node.children = [];
        this._cells.forEach(cell => {
            if (cell) this._node.children.push(cell.toObject());
        });

        return this._node;
    }

    /**
     * Initialize the row node.
     * @param {{}} node - The row node.
     * @returns {undefined}
     * @private
     */
    _init(node) {
        this._node = node;
        this._cells = [];
        this._node.children.forEach(cellNode => {
            const cell = new Cell(this, cellNode);
            this._cells[cell.columnNumber()] = cell;
        });
    }
}

module.exports = Row;

/*
<row r="6" spans="1:9" x14ac:dyDescent="0.25">
    <c r="A6" s="1" t="s">
        <v>2</v>
    </c>
    <c r="B6" s="1"/>
    <c r="C6" s="1"/>
</row>
*/
