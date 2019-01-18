"use strict";

/* eslint camelcase:off */
const RichText = require("./RichText");

class RichTexts {
    constructor(value) {
        this._node = [];
        if (value) {
            for (let i = 0; i < value.length; i++) {
                this._node.push(new RichText(value[i], null));
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
        this._node.push(new RichText(value, styles));
        return this;
    }

    toXml() {
        const node = [];
        for (let i = 0; i < this._node.length; i++) {
            node.push(this._node[i].toXml());
        }
        return node;
    }
}

// IE doesn't support function names so explicitly set it.
if (!RichTexts.name) RichTexts.name = "RichTexts";

module.exports = RichTexts;