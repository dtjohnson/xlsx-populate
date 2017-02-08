"use strict";

// TODO: Docs
// TODO: Tests

// TODO in future: style

const debug = require("./debug")("Column");

class Column {
    constructor(sheet, node) {
        debug("constructor(...)");
        this._sheet = sheet;
        this._node = node;
    }

    address() {
        // TODO
    }

    cell() {
        // TODO
    }

    columnName() {
        // TODO
    }

    columnNumber() {
        // TODO
    }

    fullAddress() {
        // TODO
    }

    sheet() {
        return this._sheet;
    }

    // TODO: xq
    width(width) {
        debug('width(%o)', arguments);
        if (arguments.length === 0) {
            return this._node.getAttribute("customWidth") === "1" ? parseFloat(this._node.getAttribute("width")) : null;
        } else if (arguments.length === 1) {
            if (width) {
                this._node.setAttribute("width", width);
                this._node.setAttribute("customWidth", "1");
            } else {
                this._node.removeAttribute("width");
                this._node.removeAttribute("customWidth");
            }

            return this;
        } else {
            throw new Error();
        }
    }

    workbook() {
        // TODO
    }
}

module.exports = Column;
