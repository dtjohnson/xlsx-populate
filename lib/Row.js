"use strict";

const Cell = require("./Cell");
const utils = require("./utils");

class Row {
    constructor(sheet, node) {
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

    search() {

    }

    replace() {

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
