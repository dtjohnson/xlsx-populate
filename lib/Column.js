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
            return this._node.attributes.customWidth ? this._node.attributes.width : undefined;
        } else if (arguments.length === 1) {
            if (width) {
                this._node.attributes.width = width;
                this._node.attributes.customWidth = 1;
            } else {
                delete this._node.attributes.width;
                delete this._node.attributes.customWidth;
            }

            return this;
        } else {
            throw new Error();
        }
    }

    workbook() {
        // TODO
    }

    toObject() {
        return this._node;
    }
}

module.exports = Column;
