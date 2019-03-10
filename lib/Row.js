"use strict";

const _ = require("lodash");
const Cell = require("./Cell");
const regexify = require("./regexify");
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
        this._sheet = sheet;
        this._init(node);
    }

    /* PUBLIC */

    /**
     * Get the address of the row.
     * @param {{}} [opts] - Options
     * @param {boolean} [opts.includeSheetName] - Include the sheet name in the address.
     * @param {boolean} [opts.anchored] - Anchor the address.
     * @returns {string} The address
     */
    address(opts) {
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
        let columnNumber = columnNameOrNumber;
        if (typeof columnNameOrNumber === 'string') {
            columnNumber = addressConverter.columnNameToNumber(columnNameOrNumber);
        }

        if (columnNumber < 1) throw new RangeError(`Invalid column number ${columnNumber}. Remember that spreadsheets use 1-based indexing.`);

        // Return an existing cell.
        if (this._cells[columnNumber]) return this._cells[columnNumber];

        // No cell exists for this.
        // Check if there is an existing row/column style for the new cell.
        let styleId;
        const rowStyleId = this._node.attributes.s;
        const columnStyleId = this.sheet().existingColumnStyleId(columnNumber);

        // Row style takes priority. If a cell has both row and column styles it should have created a cell entry with a cell-specific style.
        if (!_.isNil(rowStyleId)) styleId = rowStyleId;
        else if (!_.isNil(columnStyleId)) styleId = columnStyleId;

        // Create the new cell.
        const cell = new Cell(this, columnNumber, styleId);
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
        return this._node.attributes.r;
    }

    /**
     * Gets the parent sheet of the row.
     * @returns {Sheet} The parent sheet.
     */
    sheet() {
        return this._sheet;
    }

    /**
     * Gets an individual style.
     * @param {string} name - The name of the style.
     * @returns {*} The style.
     *//**
     * Gets multiple styles.
     * @param {Array.<string>} names - The names of the style.
     * @returns {object.<string, *>} Object whose keys are the style names and values are the styles.
     *//**
     * Sets an individual style.
     * @param {string} name - The name of the style.
     * @param {*} value - The value to set.
     * @returns {Cell} The cell.
     *//**
	 * Sets multiple styles.
	 * @param {object.<string, *>} styles - Object whose keys are the style names and values are the styles to set.
	 * @returns {Cell} The cell.
     *//**
     * Sets to a specific style
     * @param {Style} style - Style object given from stylesheet.createStyle
     * @returns {Cell} The cell.
     */
    style() {
        return new ArgHandler("Row.style")
            .case('string', name => {
                // Get single value
                this._createStyleIfNeeded();
                return this._style.style(name);
            })
            .case('array', names => {
                // Get list of values
                const values = {};
                names.forEach(name => {
                    values[name] = this.style(name);
                });

                return values;
            })
            .case(['string', '*'], (name, value) => {
                this._createCellStylesIfNeeded();

                // Style each existing cell within this row. (Cells don't inherit ow/column styles.)
                _.forEach(this._cells, cell => {
                    if (cell) cell.style(name, value);
                });

                // Set the style on the row.
                this._createStyleIfNeeded();
                this._style.style(name, value);

                return this;
            })
            .case('object', nameValues => {
                // Object of key value pairs to set
                for (const name in nameValues) {
                    if (!nameValues.hasOwnProperty(name)) continue;
                    const value = nameValues[name];
                    this.style(name, value);
                }

                return this;
            })
            .case('Style', style => {
                this._createCellStylesIfNeeded();

                // Style each existing cell within this row. (Cells don't inherit ow/column styles.)
                _.forEach(this._cells, cell => {
                    if (cell) cell.style(style);
                });

                this._style = style;
                this._node.attributes.s = style.id();
                this._node.attributes.customFormat = 1;

                return this;
            })
            .handle(arguments);
    }

    /**
     * Get the parent workbook.
     * @returns {Workbook} The parent workbook.
     */
    workbook() {
        return this.sheet().workbook();
    }

    /**
     * Append horizontal page break after the row.
     * @returns {Row} the row.
     */
    addPageBreak() {
        this.sheet().horizontalPageBreaks().add(this.rowNumber());
        return this;
    }

    /* INTERNAL */

    /**
     * Clear cells that are using a given shared formula ID.
     * @param {number} sharedFormulaId - The shared formula ID.
     * @returns {undefined}
     * @ignore
     */
    clearCellsUsingSharedFormula(sharedFormulaId) {
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
        pattern = regexify(pattern);

        const matches = [];
        this._cells.forEach(cell => {
            if (!cell) return;
            if (cell.find(pattern, replacement)) matches.push(cell);
        });

        return matches;
    }

    /**
     * Check if the row has a cell at the given column number.
     * @param {number} columnNumber - The column number.
     * @returns {boolean} True if a cell exists, false otherwise.
     * @ignore
     */
    hasCell(columnNumber) {
        if (columnNumber < 1) throw new RangeError(`Invalid column number ${columnNumber}. Remember that spreadsheets use 1-based indexing.`);
        return !!this._cells[columnNumber];
    }

    /**
     * Check if the column has a style defined.
     * @returns {boolean} True if a style exists, false otherwise.
     * @ignore
     */
    hasStyle() {
        return !_.isNil(this._node.attributes.s);
    }

    /**
     * Returns the nax used column number.
     * @returns {number} The max used column number.
     * @ignore
     */
    minUsedColumnNumber() {
        return _.findIndex(this._cells);
    }

    /**
     * Returns the nax used column number.
     * @returns {number} The max used column number.
     * @ignore
     */
    maxUsedColumnNumber() {
        return this._cells.length - 1;
    }

    /**
     * Convert the row to an object.
     * @returns {{}} The object form.
     * @ignore
     */
    toXml() {
        return this._node;
    }

    /* PRIVATE */

    /**
     * If a column node is already defined that intersects with this row and that column has a style set, we
     * need to make sure that a cell node exists at the intersection so we can style it appropriately.
     * Fetching the cell will force a new cell node to be created with a style matching the column.
     * @returns {undefined}
     * @private
     */
    _createCellStylesIfNeeded() {
        this.sheet().forEachExistingColumnNumber(columnNumber => {
            if (!_.isNil(this.sheet().existingColumnStyleId(columnNumber))) this.cell(columnNumber);
        });
    }

    /**
     * Create a style for this row if it doesn't already exist.
     * @returns {undefined}
     * @private
     */
    _createStyleIfNeeded() {
        if (!this._style) {
            const styleId = this._node.attributes.s;
            this._style = this.workbook().styleSheet().createStyle(styleId);
            this._node.attributes.s = this._style.id();
            this._node.attributes.customFormat = 1;
        }
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
        this._node.children = this._cells;
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
