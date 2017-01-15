"use strict";

const utils = require("./utils");

class Cell {
    constructor(row, node) {
        this._row = row;
        this._node = node;
    }

    /**
     * Gets the address of the cell (e.g. "A5").
     * @returns {string} The cell address.
     */
    address() {
        if (arguments.length > 0) throw new Error('Cell.address: Cannot be set.');
        return this._node.getAttribute("r");
    }

    /**
     * Gets the column number of the cell (1-based).
     * @returns {number} The column number.
     */
    columnNumber() {
        if (arguments.length > 0) throw new Error('Cell.columnNumber: Cannot be set.');
        return utils.addressToRowAndColumn(this.address()).column;
    }

    toString() {
        return this._node.toString();
    }
}

module.exports = Cell;

/*
<c r="A6" s="1" t="s">
    <v>2</v>
</c>
*/
