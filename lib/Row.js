"use strict";

const Cell = require("./Cell");

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

    rowNumber() {
        return parseInt(this._node.getAttribute("r"));
    }

    toString() {
        return this._node.toString();
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
