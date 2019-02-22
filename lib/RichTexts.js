"use strict";

/* eslint camelcase:off */
const RichText = require("./RichText");

class RichTexts {
    constructor(value, cell) {
        this._node = [];
        this._cell = cell;
        if (value) {
            for (let i = 0; i < value.length; i++) {
                this._node.push(new RichText(value[i], null, cell));
            }
        }
    }

    get(index) {
        return this._node[index];
    }

    remove(index) {
        this._node.splice(index, 1);
        return this;
    }

    add(value, styles) {
        this._node.push(new RichText(value, styles, this._cell));
        return this;
    }

    length() {
        return this._node.length;
    }

    toXml() {
        const node = [];
        for (let i = 0; i < this._node.length; i++) {
            node.push(this._node[i].toXml());
        }
        return node;
    }

    cell() {
        return this._cell;
    }
}

// IE doesn't support function names so explicitly set it.
if (!RichTexts.name) RichTexts.name = "RichTexts";

module.exports = RichTexts;
