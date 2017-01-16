"use strict";

const Style = require("./Style");

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

    iterateCells(cb) {
        for (let row = this.startCell().rowNumber(); row <= this.endCell().rowNumber(); row++) {
            for (let column = this.startCell().columnNumber(); column <= this.endCell().columnNumber(); column++) {
                cb(this.sheet().cell(row, column), row, column);
            }
        }

        return this;
    }

    // TODO: Relative cell refs

    values(values) {
        if (arguments.length === 0) {
            // Getter
            values = [];
            this.iterateCells((cell, row, column) => {
                const ri = row - this.startCell().rowNumber();
                const ci = column - this.startCell().columnNumber();
                if (!values[ri]) values[ri] = [];
                values[ri][ci] = cell.value();
            });

            return values;
        } else if (arguments.length === 1) {
            if (Array.isArray(values)) {
                this.iterateCells((cell, row, column) => {
                    const ri = row - this.startCell().rowNumber();
                    const ci = column - this.startCell().columnNumber();
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
            this.iterateCells((cell, row, column) => {
                const ri = row - this.startCell().rowNumber();
                const ci = column - this.startCell().columnNumber();
                if (!styles[ri]) styles[ri] = [];
                styles[ri][ci] = cell.style();
            });

            this._style = new Style(styles);
        }

        return this._style;
    }

    formula() {

    }

    relativeCell(rowOffset, columnOffset) {
        return this.startCell().relativeCell(rowOffset, columnOffset);
    }
}

module.exports = Range;
