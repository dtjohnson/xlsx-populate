"use strict";

// TODO: JSDoc
// TODO: Tests

const Style = require("./Style");
const debug = require("./debug")('Range');
const ArgParser = require("./ArgParser");

class Range {
    constructor(startCell, endCell) {
        this._startCell = startCell;
        this._endCell = endCell;
    }

    groupWith() {
        // TODO
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

    /*
     <mergeCells count="1">
         <mergeCell ref="A6:C6"/>
     </mergeCells>
     */
    merge() {
        // TODO
    }

    unmerge() {
        // TODO
    }

    address() {
        return `${this.startCell().address()}:${this.endCell().address()}`;
    }

    fullAddress() {
        // TODO
    }

    numRows() {
        return this.endCell().rowNumber() - this.startCell().rowNumber() + 1;
    }

    numColumns() {
        return this.endCell().columnNumber() - this.startCell().columnNumber() + 1;
    }

    // TODO: This is 1-indexed. That's the right thing to do, right?
    cell(relativeRowNumber, relativeColumnNumber) {
        const rowNumber = this.startCell().rowNumber() + relativeRowNumber - 1;
        const columnNumber = this.startCell().columnNumber() + relativeColumnNumber - 1;

        return this.sheet().cell(rowNumber, columnNumber);
    }

    forEach(handler) {
        for (let relativeRowNumber = 1; relativeRowNumber <= this.numRows(); relativeRowNumber++) {
            for (let relativeColumnNumber = 1; relativeColumnNumber <= this.numColumns(); relativeColumnNumber++) {
                handler(this.cell(relativeRowNumber, relativeColumnNumber), relativeRowNumber, relativeColumnNumber);
            }
        }

        return this;
    }

    map(handler) {
        const result = [];
        this.forEach((cell, relativeRowNumber, relativeColumnNumber) => {
            const ri = relativeRowNumber - 1;
            const ci = relativeColumnNumber - 1;
            if (!result[ri]) result[ri] = [];
            result[ri][ci] = handler(cell, relativeRowNumber, relativeColumnNumber);
        });

        return result;
    }

    reduce() {
        // TODO
    }

    // TODO: Is this really necessary?
    forEachValue(values, handler) {
        return this.forEach((cell, relativeRowNumber, relativeColumnNumber) => {
            const ri = relativeRowNumber - 1;
            const ci = relativeColumnNumber - 1;
            if (values[ri] && values[ri][ci] !== undefined) {
                handler(values[ri][ci], cell, relativeRowNumber, relativeColumnNumber);
            }
        });
    }

    value() {
        debug("value(%o)", arguments);
        return new ArgParser("Range.value")
            .case(() => {
                // Get values
                return this.map(cell => cell.value());
            })
            .case(Function, valueFunc => {
                // Set a value for the cells to the result of a function
                return this.forEach((cell, relativeRowNumber, relativeColumnNumber) => {
                    cell.value(valueFunc(cell, relativeRowNumber, relativeColumnNumber));
                });
            })
            .case(Array, values => {
                // Set value for the cells using an array of matching dimension
                return this.forEachValue(values, (value, cell) => cell.value(value));
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
            .case([String, Function], (name, valueFunc) => {
                // Set a single value for the cells to the result of a function
                return this.forEach((cell, relativeRowNumber, relativeColumnNumber) => {
                    cell.style(name, valueFunc(cell, relativeRowNumber, relativeColumnNumber));
                });
            })
            .case([String, Array], (name, values) => {
                // Set a single value for the cells using an array of matching dimension
                return this.forEachValue(values, (value, cell) => cell.style(name, value));
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

    // TODO: _ArgParser
    formula(formula) {
        debug("formula(%o)", arguments);
        if (arguments.length === 0) {

        } else if (arguments.length === 1) {
            const sharedFormulaId = ++this.startCell().sheet()._maxSharedFormulaId;
            this.forEach((cell, relativeRowNumber, relativeColumnNumber) => {
                if (relativeRowNumber === 1 && relativeColumnNumber === 1) {
                    cell._setSharedFormula(formula, sharedFormulaId, this.address());
                } else {
                    cell._setSharedFormula(null, sharedFormulaId);
                }
            });

            return this;
        } else {
            throw new Error();
        }
    }
}

module.exports = Range;
