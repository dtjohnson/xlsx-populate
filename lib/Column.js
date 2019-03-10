"use strict";

const ArgHandler = require("./ArgHandler");
const addressConverter = require('./addressConverter');

// Default column width.
const defaultColumnWidth = 9.140625;

/**
 * A column.
 */
class Column {
    // /**
    //  * Creates a new Column.
    //  * @param {Sheet} sheet - The parent sheet.
    //  * @param {{}} node - The column node.
    //  * @constructor
    //  * @ignore
    //  * @private
    //  */
    constructor(sheet, node) {
        this._sheet = sheet;
        this._node = node;
    }

    /* PUBLIC */

    /**
     * Get the address of the column.
     * @param {{}} [opts] - Options
     * @param {boolean} [opts.includeSheetName] - Include the sheet name in the address.
     * @param {boolean} [opts.anchored] - Anchor the address.
     * @returns {string} The address
     */
    address(opts) {
        return addressConverter.toAddress({
            type: 'column',
            columnName: this.columnName(),
            sheetName: opts && opts.includeSheetName && this.sheet().name(),
            columnAnchored: opts && opts.anchored
        });
    }

    /**
     * Get a cell within the column.
     * @param {number} rowNumber - The row number.
     * @returns {Cell} The cell in the column with the given row number.
     */
    cell(rowNumber) {
        return this.sheet().cell(rowNumber, this.columnNumber());
    }

    /**
     * Get the name of the column.
     * @returns {string} The column name.
     */
    columnName() {
        return addressConverter.columnNumberToName(this.columnNumber());
    }

    /**
     * Get the number of the column.
     * @returns {number} The column number.
     */
    columnNumber() {
        return this._node.attributes.min;
    }

    /**
     * Gets a value indicating whether the column is hidden.
     * @returns {boolean} A flag indicating whether the column is hidden.
     *//**
     * Sets whether the column is hidden.
     * @param {boolean} hidden - A flag indicating whether to hide the column.
     * @returns {Column} The column.
     */
    hidden() {
        return new ArgHandler("Column.hidden")
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
     * Get the parent sheet.
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
        return new ArgHandler("Column.style")
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
                // If a row node is already defined that intersects with this column and that row has a style set, we
                // need to make sure that a cell node exists at the intersection so we can style it appropriately.
                // Fetching the cell will force a new cell node to be created with a style matching the column. So we
                // will fetch and style the cell at each row that intersects this column if it is already present or it
                // has a style defined.
                this.sheet().forEachExistingRow(row => {
                    if (row.hasStyle() || row.hasCell(this.columnNumber())) {
                        row.cell(this.columnNumber()).style(name, value);
                    }
                });

                // Set a single value for all cells to a single value
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
                // See Large Comment Above
                this.sheet().forEachExistingRow(row => {
                    if (row.hasStyle() || row.hasCell(this.columnNumber())) {
                        row.cell(this.columnNumber()).style(style);
                    }
                });

                this._style = style;
                this._node.attributes.style = style.id();

                return this;
            })
            .handle(arguments);
    }

    /**
     * Gets the width.
     * @returns {undefined|number} The width (or undefined).
     *//**
     * Sets the width.
     * @param {number} width - The width of the column.
     * @returns {Column} The column.
     */
    width(width) {
        return new ArgHandler("Column.width")
            .case(() => {
                return this._node.attributes.customWidth ? this._node.attributes.width : undefined;
            })
            .case('number', width => {
                this._node.attributes.width = width;
                this._node.attributes.customWidth = 1;
                return this;
            })
            .case('nil', () => {
                delete this._node.attributes.width;
                delete this._node.attributes.customWidth;
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
     * Append vertical page break after the column.
     * @returns {Column} the column.
     */
    addPageBreak() {
        this.sheet().verticalPageBreaks().add(this.columnNumber());
        return this;
    }

    /* INTERNAL */

    /**
     * Convert the column to an XML object.
     * @returns {{}} The XML form.
     * @ignore
     */
    toXml() {
        return this._node;
    }

    /* PRIVATE */

    /**
     * Create a style for this column if it doesn't already exist.
     * @returns {undefined}
     * @private
     */
    _createStyleIfNeeded() {
        if (!this._style) {
            const styleId = this._node.attributes.style;
            this._style = this.workbook().styleSheet().createStyle(styleId);
            this._node.attributes.style = this._style.id();

            if (!this.width()) this.width(defaultColumnWidth);
        }
    }
}

module.exports = Column;
