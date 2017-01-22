"use strict";

const Style = require("./Style");
const debug = require("./debug")('Range');
const _ArgParser = require("./_ArgParser");

class Range {
    constructor(startCell, endCell) {
        this._startCell = startCell;
        this._endCell = endCell;
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

    }

    unmerge() {

    }

    address() {
        return `${this.startCell().address()}:${this.endCell().address()}`;
    }

    numRows() {
        return this.endCell().rowNumber() - this.startCell().rowNumber() + 1;
    }

    numColumns() {
        return this.endCell().columnNumber() - this.startCell().columnNumber() + 1;
    }

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

    forEachValue(values, handler) {
        this.forEach((cell, relativeRowNumber, relativeColumnNumber) => {
            const ri = relativeRowNumber - 1;
            const ci = relativeColumnNumber - 1;
            if (values[ri] && values[ri][ci] !== undefined) {
                handler(values[ri][ci], cell, relativeRowNumber, relativeColumnNumber);
            }
        });
    }

    value(values) {
        if (arguments.length === 0) {
            // Getter
            values = [];
            this.forEach((cell, relativeRowNumber, relativeColumnNumber) => {
                const ri = relativeRowNumber - 1;
                const ci = relativeColumnNumber - 1;
                if (!values[ri]) values[ri] = [];
                values[ri][ci] = cell.value();
            });

            return values;
        } else if (arguments.length === 1) {
            if (Array.isArray(values)) {
                this.forEach((cell, relativeRowNumber, relativeColumnNumber) => {
                    const ri = relativeRowNumber - 1;
                    const ci = relativeColumnNumber - 1;
                    cell.value(values[ri][ci]);
                });
            } else if (typeof values === "function") {
                this.forEach(cell => cell.value(values(cell)));
            } else {
                this.forEach(cell => cell.value(values));
            }

            return this;
        } else {
            throw new Error("Range.values: Unexpected number of arguments");
        }
    }

    /**
     *
     * @returns {Style}
     */
    style() {
        debug("style(%o)", arguments);
        return new _ArgParser("Cell.style")
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
                this.forEach((cell, relativeRowNumber, relativeColumnNumber) => {
                    cell.style(name, valueFunc(cell, relativeRowNumber, relativeColumnNumber));
                });

                return this;
            })
            .case([String, Array], (name, values) => {
                // Set a single value for the cells using an array of matching dimension
                this.forEachValue(values, (value, cell) => cell.style(name, value));
                return this;
            })
            .case([String, undefined], (name, value) => {
                // Set a single value for all cells to a single value
                this.forEach(cell => cell.style(name, value));
                return this;
            })
            .case(Object, nameValues => {
                // Object of key value pairs to set
                for (const name in nameValues) {
                    if (!nameValues.hasOwnProperty(name)) continue;
                    const value = nameValues[name];
                    this.style(name, value);
                }
            })
            .parse(arguments);
    }

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

    relativeCell(rowOffset, columnOffset) {
        return this.startCell().relativeCell(rowOffset, columnOffset);
    }
}

module.exports = Range;
