"use strict";

const Style = require("./Style");
const debug = require("./debug")('Range');

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

    iterateCells(cb) {
        for (let relativeRowNumber = 1; relativeRowNumber <= this.numRows(); relativeRowNumber++) {
            for (let relativeColumnNumber = 1; relativeColumnNumber <= this.numColumns(); relativeColumnNumber++) {
                cb(this.cell(relativeRowNumber, relativeColumnNumber), relativeRowNumber, relativeColumnNumber);
            }
        }

        return this;
    }

    // TODO: Relative cell refs

    values(values) {
        if (arguments.length === 0) {
            // Getter
            values = [];
            this.iterateCells((cell, relativeRowNumber, relativeColumnNumber) => {
                const ri = relativeRowNumber - 1;
                const ci = relativeColumnNumber - 1;
                if (!values[ri]) values[ri] = [];
                values[ri][ci] = cell.value();
            });

            return values;
        } else if (arguments.length === 1) {
            if (Array.isArray(values)) {
                this.iterateCells((cell, relativeRowNumber, relativeColumnNumber) => {
                    const ri = relativeRowNumber - 1;
                    const ci = relativeColumnNumber - 1;
                    cell.value(values[ri][ci]);
                });
            } else {
                this.iterateCells(cell => cell.value(values));
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
        if (!this._style) {
            const styles = [];
            this.iterateCells((cell, relativeRowNumber, relativeColumnNumber) => {
                const ri = relativeRowNumber - 1;
                const ci = relativeColumnNumber - 1;
                if (!styles[ri]) styles[ri] = [];
                styles[ri][ci] = cell.style();
            });

            this._style = new Style(styles);
        }

        return this._style;
    }

    formula(formula) {
        debug("formula(%o)", arguments);
        if (arguments.length === 0) {

        } else if (arguments.length === 1) {
            const sharedFormulaId = ++this.startCell().sheet()._maxSharedFormulaId;
            this.iterateCells((cell, relativeRowNumber, relativeColumnNumber) => {
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
