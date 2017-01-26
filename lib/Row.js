"use strict";

// TODO: Tests
// TODO: JSDoc

// TODO in future: style

const Cell = require("./Cell");
const utils = require("./utils");
const debug = require("./debug")('Row');

class Row {
    constructor(sheet, node) {
        debug("constructor(...)");

        this._sheet = sheet;
        this._node = node;

        this._cells = [];
        const cellNodes = this._node.childNodes;
        for (let i = 0; i < cellNodes.length; i++) {
            const cellNode = cellNodes[i];
            const cell = new Cell(this, cellNode);
            this._cells[cell.columnNumber()] = cell;
        }
    }

    sheet() {
        return this._sheet;
    }

    workbook() {
        return this.sheet().workbook();
    }

    rowNumber() {
        return parseInt(this._node.getAttribute("r"));
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
        debug('height(%o)', arguments)
        if (arguments.length === 0) {
            return this._node.getAttribute("customHeight") === "1" ? parseFloat(this._node.getAttribute("ht")) : null;
        } else if (arguments.length === 1) {
            if (height) {
                this._node.setAttribute("ht", height);
                this._node.setAttribute("customHeight", "1");
            } else {
                this._node.removeAttribute("ht");
                this._node.removeAttribute("customHeight");
            }

            return this;
        } else {
            throw new Error();
        }
    }

    cell(columnNumber) {
        if (this._cells[columnNumber]) return this._cells[columnNumber];
        const address = utils.rowAndColumnToAddress(this.rowNumber(), columnNumber);
        const cellNode = this._node.ownerDocument.createElement("c");
        cellNode.setAttribute("r", address);
        const cell = new Cell(this, cellNode);
        this._cells[columnNumber] = cell;
        return cell;
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
