"use strict";

// TODO: JSDoc
// TODO: Tests

const Style = require("./Style");
const debug = require("./debug")('Range');
const ArgParser = require("./ArgParser");
const addressConverter = require("./addressConverter");

class Range {
    constructor(startCell, endCell) {
        this._startCell = startCell;
        this._endCell = endCell;
    }

    groupWith() {
        return this.workbook().group(this, arguments);
    }

    sheet() {
        return this.startCell().sheet();
    }

    workbook() {
        return this.sheet().workbook();
    }

    startCell() {
        return this._startCell;
    }

    endCell() {
        return this._endCell;
    }

    merged(merged) {
        if (arguments.length) {
            if (merged) this.sheet().mergeCells(this.address());
            else this.sheet().unergedCells(this.address());
            return this;
        }

        return this.sheet().areCellsMerged(this.address());
    }

    address(opts) {
        return addressConverter.toAddress({
            type: 'range',
            startRowNumber: this.startCell().rowNumber(),
            startRowAnchored: opts.startRowAnchored,
            startColumnName: this.startCell().columnName(),
            startColumnAnchored: opts.startColumnAnchored,
            endRowNumber: this.endCell().rowNumber(),
            endRowAnchored: opts.endRowAnchored,
            endColumnName: this.endCell().columnName(),
            endColumnAnchored: opts.endColumnAnchored,
            sheetName: opts && opts.includeSheetName && this.sheet().name()
        });
    }

    numRows() {
        return this.endCell().rowNumber() - this.startCell().rowNumber() + 1;
    }

    numColumns() {
        return this.endCell().columnNumber() - this.startCell().columnNumber() + 1;
    }

    relativeCell(ri, ci) {
        return this.startCell().relativeCell(ri, ci);
    }

    forEach(callback) {
        for (let ri = 0; ri < this.numRows(); ri++) {
            for (let ci = 0; ci < this.numColumns(); ci++) {
                callback(this.relativeCell(ri, ci), ri, ci, this);
            }
        }

        return this;
    }

    map(callback) {
        const result = [];
        this.forEach((cell, ri, ci) => {
            if (!result[ri]) result[ri] = [];
            result[ri][ci] = callback(cell, ri, ci, this);
        });

        return result;
    }

    reduce(callback, initialValue) {
        let accumulator = initialValue;
        this.forEach((cell, ri, ci) => {
            accumulator = callback(accumulator, cell, ri, ci, this);
        });

        return accumulator;
    }

    tap(callback) {
        callback(this);
        return this;
    }

    thru(callback) {
        return callback(this);
    }

    value() {
        debug("value(%o)", arguments);
        return new ArgParser("Range.value")
            .case(() => {
                // Get values
                return this.map(cell => cell.value());
            })
            .case(Function, callback => {
                // Set a value for the cells to the result of a function
                return this.forEach((cell, ri, ci) => {
                    cell.value(callback(cell, ri, ci, this));
                });
            })
            .case(Array, values => {
                // Set value for the cells using an array of matching dimension
                return this.forEach((cell, ri, ci) => {
                    if (values[ri] && values[ri][ci] !== undefined) {
                        cell.value(values[ri][ci]);
                    }
                });
            })
            .case(undefined, value => {
                // Set the value for all cells to a single value
                return this.forEach(cell => cell.value(value));
            })
            .parse(arguments);
    }

    /**
     *
     * @returns {Style}
     */
    style() {
        debug("style(%o)", arguments);
        return new ArgParser("Range.style")
            .case(String, name => {
                // Get single value
                return this.map(cell => cell.style(name));
            })
            .case(Array, names => {
                // Get list of values
                const values = {};
                names.forEach(name => {
                    values[name] = this.style(name);
                });

                return values;
            })
            .case([String, Function], (name, callback) => {
                // Set a single value for the cells to the result of a function
                return this.forEach((cell, ri, ci) => {
                    cell.style(name, callback(cell, ri, ci, this));
                });
            })
            .case([String, Array], (name, values) => {
                // Set a single value for the cells using an array of matching dimension
                return this.forEach((cell, ri, ci) => {
                    if (values[ri] && values[ri][ci] !== undefined) {
                        cell.style(name, values[ri][ci]);
                    }
                });
            })
            .case([String, undefined], (name, value) => {
                // Set a single value for all cells to a single value
                return this.forEach(cell => cell.style(name, value));
            })
            .case(Object, nameValues => {
                // Object of key value pairs to set
                for (const name in nameValues) {
                    if (!nameValues.hasOwnProperty(name)) continue;
                    const value = nameValues[name];
                    this.style(name, value);
                }

                return this;
            })
            .parse(arguments);
    }

    clear() {
        return this.value(undefined);
    }

    formula(formula) {
        debug("formula(%o)", arguments);
        return new ArgParser("Range.formula")
            .case(() => {
                // TODO: What if not shared?
                return this.startCell().formula();
            })
            .case(String, formula => {
                // TODO: Switch to some better method instead of private field.
                const sharedFormulaId = ++this.startCell().sheet()._maxSharedFormulaId;
                this.forEach((cell, ri, ci) => {
                    if (ri === 0 && ci === 0) {
                        cell.setSharedFormula(formula, sharedFormulaId, this.address());
                    } else {
                        cell.setSharedFormula(null, sharedFormulaId);
                    }
                });

                return this;
            })
            .parse(arguments);
    }
}

module.exports = Range;
