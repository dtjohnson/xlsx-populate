"use strict";

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

    values(values) {
        if (arguments.length === 0) {
            // Getter
            // TODO
        } else if (arguments.length === 1) {
            if (Array.isArray(values)) {

            } else {

            }
        } else {
            throw new Error("Range.values: Unexpected number of arguments");
        }
    }

    formula() {

    }

    relativeCell(rowOffset, columnOffset) {
        return this.startCell().relativeCell(rowOffset, columnOffset);
    }
}

module.exports = Range;
