"use strict";

// TODO: Tests
// TODO: JSDoc

// TODO in future: style

const Cell = require("./Cell");
const utils = require("./utils");
const debug = require("./debug")('Row');

/**
 * A row.
 */
class Row {
    constructor(sheet, node) {
        debug("constructor(...)");
        this._sheet = sheet;
        this._node = node;

        this._cells = [];
        this._node.children.forEach(cellNode => {
            const cell = new Cell(this, cellNode);
            this._cells[cell.columnNumber()] = cell;
        });
    }

    sheet() {
        return this._sheet;
    }

    workbook() {
        return this.sheet().workbook();
    }

    rowNumber() {
        return this._node.attributes.r;
    }

    address() {
        // TODO
    }

    fullAddress() {
        // TODO
    }

    // TODO: Make private?
    find(pattern) {
        pattern = utils.getRegExpForSearch(pattern);

        const matches = [];
        this._cells.forEach(cell => {
            if (!cell) return;
            if (cell.find(pattern)) matches.push(cell);
        });

        return matches;
    }

    // TODO: Make private?
    replace(pattern, replacement) {
        pattern = utils.getRegExpForSearch(pattern);

        let count = 0;
        this._cells.forEach(cell => {
            if (!cell) return;
            if (cell.replace(pattern, replacement)) count++;
        });

        return count;
    }

    // TODO: xq
    height(height) {
        debug('height(%o)', arguments);
        if (arguments.length === 0) {
            return this._node.attributes.customHeight ? this._node.attributes.ht : undefined;
        } else if (arguments.length === 1) {
            if (height) {
                this._node.attributes.ht = height;
                this._node.attributes.customHeight = 1;
            } else {
                delete this._node.attributes.ht;
                delete this._node.attributes.customHeight;
            }

            return this;
        } else {
            throw new Error();
        }
    }

    cell(columnNumber) {
        if (this._cells[columnNumber]) return this._cells[columnNumber];
        const address = utils.rowAndColumnToAddress(this.rowNumber(), columnNumber);
        const cellNode = { name: 'c', attributes: { r: address }, children: [] };
        const cell = new Cell(this, cellNode);
        this._cells[columnNumber] = cell;
        return cell;
    }

    toObject() {
        // Cells must be in order.
        this._node.children = [];
        this._cells.forEach(cell => {
            if (cell) this._node.children.push(cell.toObject());
        });

        return this._node;
    }
}

module.exports = Row;

/*
<row r="6" spans="1:9" x14ac:dyDescent="0.25">
    <c r="A6" s="1" t="s">
        <v>2</v>
    </c>
    <c r="B6" s="1"/>
    <c r="C6" s="1"/>
</row>
*/
