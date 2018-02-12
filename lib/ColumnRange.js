"use strict";

const ArgHandler = require("./ArgHandler");
const addressConverter = require('./addressConverter');

class ColumnRange {
    constructor(startColumn, endColumn) {
        this._startColumn = startColumn;
        this._endColumn = endColumn;
        this._findRangeExtent();
    }

    /* PUBLIC */

    address(opts) {
        return addressConverter.toAddress({
            type: 'columnRange',
            startColumnName: this.startColumn().columnName(),
            startColumnAnchored: opts && (opts.startColumnAnchored || opts.anchored),
            endColumnName: this.endColumn().columnName(),
            endColumnAnchored: opts && (opts.endColumnAnchored || opts.anchored),
            sheetName: opts && opts.includeSheetName && this.sheet().name()
        });
    }

    column(ci) {
        return this.sheet().column(this._minColumnNumber + ci);
    }

    columns() {
        return this.map(column => column);
    }

    endColumn() {
        return this._endColumn;
    }

    forEach(callback) {
        for (let ci = 0; ci < this._numColumns; ci++) {
            callback(this.column(ci), ci, this);
        }

        return this;
    }

    hidden() {
        return this._getSet('hidden', arguments);
    }

    map(callback) {
        const result = [];
        this.forEach((cell, ci) => {
            result[ci] = callback(cell, ci, this);
        });

        return result;
    }

    reduce(callback, initialValue) {
        let accumulator = initialValue;
        this.forEach((cell, ci) => {
            accumulator = callback(accumulator, cell, ci, this);
        });

        return accumulator;
    }

    sheet() {
        return this.startColumn().sheet();
    }

    startColumn() {
        return this._startColumn;
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

    width() {
        return this._getSet('width', arguments);
    }

    unionWith() {

    }

    workbook() {
        return this.sheet().workbook();
    }

    /* PRIVATE */

    _getSet(method, args) {
        return new ArgHandler(`ColumnRange.${method}`)
            .case(() => {
                // Get values
                return this.map(column => column[method]());
            })
            .case('function', callback => {
                // Set a value for the columns to the result of a function
                return this.forEach((column, ci) => {
                    column[method](callback(column, ci, this));
                });
            })
            .case('array', values => {
                // Set values for the columns using an array of matching dimension
                return this.forEach((column, ci) => {
                    if (values[ci] !== undefined) {
                        column[method](values[ci]);
                    }
                });
            })
            .case('*', value => {
                // Set the value for all columns to a single value
                return this.forEach(column => column[method](value));
            })
            .handle(args);
    }

    _findRangeExtent() {
        this._minColumnNumber = Math.min(this._startColumn.columnNumber(), this._endColumn.columnNumber());
        this._maxColumnNumber = Math.max(this._startColumn.columnNumber(), this._endColumn.columnNumber());
        this._numColumns = this._maxColumnNumber - this._minColumnNumber + 1;
    }
}

module.exports = ColumnRange;
