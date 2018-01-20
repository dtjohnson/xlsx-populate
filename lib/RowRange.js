"use strict";

const ArgHandler = require("./ArgHandler");
const addressConverter = require('./addressConverter');

class RowRange {
    constructor(startRow, endRow) {
        this._startRow = startRow;
        this._endRow = endRow;
        this._findRangeExtent();
    }

    address(opts) {
        return addressConverter.toAddress({
            type: 'rowRange',
            startRowNumber: this.startRow().rowNumber(),
            startRowAnchored: opts && (opts.startRowAnchored || opts.anchored),
            endRowNumber: this.endRow().rowNumber(),
            endRowAnchored: opts && (opts.endRowAnchored || opts.anchored),
            sheetName: opts && opts.includeSheetName && this.sheet().name()
        });
    }

    endRow() {
        return this._endRow;
    }

    forEach(callback) {
        for (let ri = 0; ri < this._numRows; ri++) {
            callback(this.row(ri), ri, this);
        }

        return this;
    }

    height() {
        return this._getSet('height', arguments);
    }

    hidden() {
        return this._getSet('hidden', arguments);
    }

    map(callback) {
        const result = [];
        this.forEach((row, ri) => {
            result[ri] = callback(row, ri, this);
        });

        return result;
    }

    reduce(callback, initialValue) {
        let accumulator = initialValue;
        this.forEach((row, ri) => {
            accumulator = callback(accumulator, row, ri, this);
        });

        return accumulator;
    }

    row(ri) {
        return this.sheet().row(this._minRowNumber + ri);
    }

    rows() {
        return this.map(row => row);
    }

    sheet() {
        return this.startRow().sheet();
    }

    startRow() {
        return this._startRow;
    }

    style() {

    }

    tap(callback) {
        callback(this);
        return this;
    }

    thru(callback) {
        return callback(this);
    }

    unionWith() {

    }

    workbook() {
        return this.sheet().workbook();
    }

    /* PRIVATE */

    _getSet(method, args) {
        return new ArgHandler(`RowRange.${method}`)
            .case(() => {
                // Get values
                return this.map(row => row[method]());
            })
            .case('function', callback => {
                // Set a value for the rows to the result of a function
                return this.forEach((row, ri) => {
                    row[method](callback(row, ri, this));
                });
            })
            .case('array', values => {
                // Set values for the rows using an array of matching dimension
                return this.forEach((row, ri) => {
                    if (values[ri] !== undefined) {
                        row[method](values[ri]);
                    }
                });
            })
            .case('*', value => {
                // Set the value for all rows to a single value
                return this.forEach(row => row[method](value));
            })
            .handle(args);
    }

    _findRangeExtent() {
        this._minRowNumber = Math.min(this._startRow.rowNumber(), this._endRow.rowNumber());
        this._maxRowNumber = Math.max(this._startRow.rowNumber(), this._endRow.rowNumber());
        this._numRows = this._maxRowNumber - this._minRowNumber + 1;
    }
}

module.exports = RowRange;
