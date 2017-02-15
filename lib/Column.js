"use strict";

// TODO in future release: style, column ranges

const debug = require("./debug")("Column");
const ArgParser = require("./ArgParser");
const addressConverter = require('./addressConverter');

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
        debug("constructor(...)");
        this._sheet = sheet;
        this._node = node;
    }

    /**
     * Get the address of the column.
     * @param {{}} [opts] - Options
     * @param {boolean} [opts.includeSheetName] - Include the sheet name in the address.
     * @param {boolean} [opts.anchored] - Anchor the address.
     * @returns {string} The address
     */
    address(opts) {
        debug("address(%o)", arguments);
        return addressConverter.toAddress({
            type: 'column',
            columnName: this.columnName(),
            sheetName: opts && opts.includeSheetName && this.sheet().name(),
            anchored: opts && opts.anchored
        });
    }

    /**
     * Get a cell within the column.
     * @param {number} rowNumber - The row number.
     * @returns {Cell} The cell in the column with the given row number.
     */
    cell(rowNumber) {
        debug("cell(%o)", arguments);
        return this.sheet().cell(rowNumber, this.columnNumber());
    }

    /**
     * Get the name of the column.
     * @returns {string} The column name.
     */
    columnName() {
        debug("columnName(%o)", arguments);
        return addressConverter.columnNumberToName(this.columnNumber());
    }

    /**
     * Get the number of the column.
     * @returns {number} The column number.
     */
    columnNumber() {
        debug("columnNumber(%o)", arguments);
        return this._node.attributes.min;
    }

    /**
     * Gets or sets whether the column is hidden.
     * @param {boolean} [hidden] - A flag indicating whether to hide the column.
     * @returns {boolean|Column} A flag indicating whether the column is hidden if getting, the column if setting.
     */
    hidden() {
        debug('hidden(%o)', arguments);
        return new ArgParser("Column.hidden")
            .case(() => {
                return this._node.attributes.hidden === 1;
            })
            .case(Boolean, hidden => {
                if (hidden) this._node.attributes.hidden = 1;
                else delete this._node.attributes.hidden;
                return this;
            })
            .parse(arguments);
    }

    /**
     * Get the parent sheet.
     * @returns {Sheet} The parent sheet.
     */
    sheet() {
        debug("sheet(%o)", arguments);
        return this._sheet;
    }

    /**
     * Convert the column to an object.
     * @returns {{}} The object form.
     * @ignore
     */
    toObject() {
        debug("toObject(%o)", arguments);
        return this._node;
    }

    /**
     * Gets or sets the width.
     * @param {number} [width] - The width of the column.
     * @returns {undefined|number|Column} The width (or undefined) if getting, the column if setting.
     */
    width(width) {
        debug('width(%o)', arguments);
        return new ArgParser("Column.width")
            .case(() => {
                return this._node.attributes.customWidth ? this._node.attributes.width : undefined;
            })
            .case(Number, width => {
                this._node.attributes.width = width;
                this._node.attributes.customWidth = 1;
                return this;
            })
            .case(undefined, () => {
                delete this._node.attributes.width;
                delete this._node.attributes.customWidth;
                return this;
            })
            .parse(arguments);
    }

    /**
     * Get the parent workbook.
     * @returns {Workbook} The parent workbook.
     */
    workbook() {
        debug("workbook(%o)", arguments);
        return this.sheet().workbook();
    }
}

module.exports = Column;
