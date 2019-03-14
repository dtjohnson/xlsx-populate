"use strict";

const ArgHandler = require("./ArgHandler");
const addressConverter = require("./addressConverter");

/**
 * A range of cells.
 */
class Range {
    // /**
    //  * Creates a new instance of Range.
    //  * @param {Cell} startCell - The start cell.
    //  * @param {Cell} endCell - The end cell.
    //  */
    constructor(startCell, endCell) {
        this._startCell = startCell;
        this._endCell = endCell;
        this._findRangeExtent(startCell, endCell);
    }

    /**
     * Get the address of the range.
     * @param {{}} [opts] - Options
     * @param {boolean} [opts.includeSheetName] - Include the sheet name in the address.
     * @param {boolean} [opts.startRowAnchored] - Anchor the start row.
     * @param {boolean} [opts.startColumnAnchored] - Anchor the start column.
     * @param {boolean} [opts.endRowAnchored] - Anchor the end row.
     * @param {boolean} [opts.endColumnAnchored] - Anchor the end column.
     * @param {boolean} [opts.anchored] - Anchor all row and columns.
     * @returns {string} The address.
     */
    address(opts) {
        return addressConverter.toAddress({
            type: 'range',
            startRowNumber: this.startCell().rowNumber(),
            startRowAnchored: opts && (opts.startRowAnchored || opts.anchored),
            startColumnName: this.startCell().columnName(),
            startColumnAnchored: opts && (opts.startColumnAnchored || opts.anchored),
            endRowNumber: this.endCell().rowNumber(),
            endRowAnchored: opts && (opts.endRowAnchored || opts.anchored),
            endColumnName: this.endCell().columnName(),
            endColumnAnchored: opts && (opts.endColumnAnchored || opts.anchored),
            sheetName: opts && opts.includeSheetName && this.sheet().name()
        });
    }

    /**
     * Gets a cell within the range.
     * @param {number} ri - Row index relative to the top-left corner of the range (0-based).
     * @param {number} ci - Column index relative to the top-left corner of the range (0-based).
     * @returns {Cell} The cell.
     */
    cell(ri, ci) {
        return this.sheet().cell(this._minRowNumber + ri, this._minColumnNumber + ci);
    }

    /**
     * Sets sheet autoFilter to this range.
     * @returns {Range} This range.
     */
    autoFilter() {
        this.sheet().autoFilter(this);

        return this;
    }

    /**
     * Get the cells in the range as a 2D array.
     * @returns {Array.<Array.<Cell>>} The cells.
     */
    cells() {
        return this.map(cell => cell);
    }

    /**
     * Clear the contents of all the cells in the range.
     * @returns {Range} The range.
     */
    clear() {
        return this.value(undefined);
    }

    /**
     * Get the end cell of the range.
     * @returns {Cell} The end cell.
     */
    endCell() {
        return this._endCell;
    }

    /**
     * Callback used by forEach.
     * @callback Range~forEachCallback
     * @param {Cell} cell - The cell.
     * @param {number} ri - The relative row index.
     * @param {number} ci - The relative column index.
     * @param {Range} range - The range.
     * @returns {undefined}
     */
    /**
     * Call a function for each cell in the range. Goes by row then column.
     * @param {Range~forEachCallback} callback - Function called for each cell in the range.
     * @returns {Range} The range.
     */
    forEach(callback) {
        for (let ri = 0; ri < this._numRows; ri++) {
            for (let ci = 0; ci < this._numColumns; ci++) {
                callback(this.cell(ri, ci), ri, ci, this);
            }
        }

        return this;
    }

    /**
     * Gets the shared formula in the start cell (assuming it's the source of the shared formula).
     * @returns {string|undefined} The shared formula.
     *//**
     * Sets the shared formula in the range. The formula will be translated for each cell.
     * @param {string} formula - The formula to set.
     * @returns {Range} The range.
     */
    formula() {
        return new ArgHandler("Range.formula")
            .case(() => {
                return this.startCell().getSharedRefFormula();
            })
            .case('string', formula => {
                const sharedFormulaId = this.sheet().incrementMaxSharedFormulaId();
                this.forEach((cell, ri, ci) => {
                    if (ri === 0 && ci === 0) {
                        cell.setSharedFormula(sharedFormulaId, formula, this.address());
                    } else {
                        cell.setSharedFormula(sharedFormulaId);
                    }
                });

                return this;
            })
            .handle(arguments);
    }

    /**
     * Callback used by map.
     * @callback Range~mapCallback
     * @param {Cell} cell - The cell.
     * @param {number} ri - The relative row index.
     * @param {number} ci - The relative column index.
     * @param {Range} range - The range.
     * @returns {*} The value to map to.
     */
    /**
     * Creates a 2D array of values by running each cell through a callback.
     * @param {Range~mapCallback} callback - Function called for each cell in the range.
     * @returns {Array.<Array.<*>>} The 2D array of return values.
     */
    map(callback) {
        const result = [];
        this.forEach((cell, ri, ci) => {
            if (!result[ri]) result[ri] = [];
            result[ri][ci] = callback(cell, ri, ci, this);
        });

        return result;
    }

    /**
     * Gets a value indicating whether the cells in the range are merged.
     * @returns {boolean} The value.
     *//**
     * Sets a value indicating whether the cells in the range should be merged.
     * @param {boolean} merged - True to merge, false to unmerge.
     * @returns {Range} The range.
     */
    merged(merged) {
        return new ArgHandler('Range.merged')
            .case(() => {
                return this.sheet().merged(this.address());
            })
            .case('*', merged => {
                this.sheet().merged(this.address(), merged);
                return this;
            })
            .handle(arguments);
    }

    /**
     * Gets the data validation object attached to the Range.
     * @returns {object|undefined} The data validation object or undefined if not set.
     *//**
     * Set or clear the data validation object of the entire range.
     * @param {object|undefined} dataValidation - Object or null to clear.
     * @returns {Range} The range.
     */
    dataValidation() {
        return new ArgHandler('Range.dataValidation')
            .case(() => {
                return this.sheet().dataValidation(this.address());
            })
            .case('boolean', obj => {
                return this.sheet().dataValidation(this.address(), obj);
            })
            .case('*', obj => {
                this.sheet().dataValidation(this.address(), obj);
                return this;
            })
            .handle(arguments);
    }

    /**
     * Callback used by reduce.
     * @callback Range~reduceCallback
     * @param {*} accumulator - The accumulated value.
     * @param {Cell} cell - The cell.
     * @param {number} ri - The relative row index.
     * @param {number} ci - The relative column index.
     * @param {Range} range - The range.
     * @returns {*} The value to map to.
     */
    /**
     * Reduces the range to a single value accumulated from the result of a function called for each cell.
     * @param {Range~reduceCallback} callback - Function called for each cell in the range.
     * @param {*} [initialValue] - The initial value.
     * @returns {*} The accumulated value.
     */
    reduce(callback, initialValue) {
        let accumulator = initialValue;
        this.forEach((cell, ri, ci) => {
            accumulator = callback(accumulator, cell, ri, ci, this);
        });

        return accumulator;
    }

    /**
     * Gets the parent sheet of the range.
     * @returns {Sheet} The parent sheet.
     */
    sheet() {
        return this.startCell().sheet();
    }

    /**
     * Gets the start cell of the range.
     * @returns {Cell} The start cell.
     */
    startCell() {
        return this._startCell;
    }

    /**
     * Gets a single style for each cell.
     * @param {string} name - The name of the style.
     * @returns {Array.<Array.<*>>} 2D array of style values.
     *//**
     * Gets multiple styles for each cell.
     * @param {Array.<string>} names - The names of the styles.
     * @returns {Object.<string, Array.<Array.<*>>>} Object whose keys are style names and values are 2D arrays of style values.
     *//**
     * Set the style in each cell to the result of a function called for each.
     * @param {string} name - The name of the style.
     * @param {Range~mapCallback} callback - The callback to provide value for the cell.
     * @returns {Range} The range.
     *//**
     * Sets the style in each cell to the corresponding value in the given 2D array of values.
     * @param {string} name - The name of the style.
     * @param {Array.<Array.<*>>} values - The style values to set.
     * @returns {Range} The range.
     *//**
     * Set the style of all cells in the range to a single style value.
     * @param {string} name - The name of the style.
     * @param {*} value - The value to set.
     * @returns {Range} The range.
     *//**
     * Set multiple styles for the cells in the range.
     * @param {object.<string,Range~mapCallback|Array.<Array.<*>>|*>} styles - Object whose keys are style names and values are either function callbacks, 2D arrays of style values, or a single value for all the cells.
     * @returns {Range} The range.
     *//**
	 * Sets to a specific style
	 * @param {Style} style - Style object given from stylesheet.createStyle
	 * @returns {Range} The range.
	 */
    style() {
        return new ArgHandler("Range.style")
            .case('string', name => {
                // Get single value
                return this.map(cell => cell.style(name));
            })
            .case('array', names => {
                // Get list of values
                const values = {};
                names.forEach(name => {
                    values[name] = this.style(name);
                });

                return values;
            })
            .case(['string', 'function'], (name, callback) => {
                // Set a single value for the cells to the result of a function
                return this.forEach((cell, ri, ci) => {
                    cell.style(name, callback(cell, ri, ci, this));
                });
            })
            .case(['string', 'array'], (name, values) => {
                // Set a single value for the cells using an array of matching dimension
                return this.forEach((cell, ri, ci) => {
                    if (values[ri] && values[ri][ci] !== undefined) {
                        cell.style(name, values[ri][ci]);
                    }
                });
            })
            .case(['string', '*'], (name, value) => {
                // Set a single value for all cells to a single value
                return this.forEach(cell => cell.style(name, value));
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
                this._style = style;
                return this.forEach(cell => cell.style(style));
            })
            .handle(arguments);
    }

    /**
     * Callback used by tap.
     * @callback Range~tapCallback
     * @param {Range} range - The range.
     * @returns {undefined}
     */
    /**
     * Invoke a callback on the range and return the range. Useful for method chaining.
     * @param {Range~tapCallback} callback - The callback function.
     * @returns {Range} The range.
     */
    tap(callback) {
        callback(this);
        return this;
    }

    /**
     * Callback used by thru.
     * @callback Range~thruCallback
     * @param {Range} range - The range.
     * @returns {*} The value to return from thru.
     */
    /**
     * Invoke a callback on the range and return the value provided by the callback. Useful for method chaining.
     * @param {Range~thruCallback} callback - The callback function.
     * @returns {*} The return value of the callback.
     */
    thru(callback) {
        return callback(this);
    }

    /**
     * Get the values of each cell in the range as a 2D array.
     * @returns {Array.<Array.<*>>} The values.
     *//**
     * Set the values in each cell to the result of a function called for each.
     * @param {Range~mapCallback} callback - The callback to provide value for the cell.
     * @returns {Range} The range.
     *//**
     * Sets the value in each cell to the corresponding value in the given 2D array of values.
     * @param {Array.<Array.<*>>} values - The values to set.
     * @returns {Range} The range.
     *//**
     * Set the value of all cells in the range to a single value.
     * @param {*} value - The value to set.
     * @returns {Range} The range.
     */
    value() {
        return new ArgHandler("Range.value")
            .case(() => {
                // Get values
                return this.map(cell => cell.value());
            })
            .case('function', callback => {
                // Set a value for the cells to the result of a function
                return this.forEach((cell, ri, ci) => {
                    cell.value(callback(cell, ri, ci, this));
                });
            })
            .case('array', values => {
                // Set value for the cells using an array of matching dimension
                return this.forEach((cell, ri, ci) => {
                    if (values[ri] && values[ri][ci] !== undefined) {
                        cell.value(values[ri][ci]);
                    }
                });
            })
            .case('*', value => {
                // Set the value for all cells to a single value
                return this.forEach(cell => cell.value(value));
            })
            .handle(arguments);
    }

    /**
     * Gets the parent workbook.
     * @returns {Workbook} The parent workbook.
     */
    workbook() {
        return this.sheet().workbook();
    }

    /**
     * Find the extent of the range.
     * @returns {undefined}
     * @private
     */
    _findRangeExtent() {
        this._minRowNumber = Math.min(this._startCell.rowNumber(), this._endCell.rowNumber());
        this._maxRowNumber = Math.max(this._startCell.rowNumber(), this._endCell.rowNumber());
        this._minColumnNumber = Math.min(this._startCell.columnNumber(), this._endCell.columnNumber());
        this._maxColumnNumber = Math.max(this._startCell.columnNumber(), this._endCell.columnNumber());
        this._numRows = this._maxRowNumber - this._minRowNumber + 1;
        this._numColumns = this._maxColumnNumber - this._minColumnNumber + 1;
    }
}

module.exports = Range;
