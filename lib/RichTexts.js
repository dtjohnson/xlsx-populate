"use strict";

/* eslint camelcase:off */
const RichTextFragment = require("./RichTextFragment");

/**
 * A RichTexts class that contains many {@link RichTextFragment}.
 */
class RichTexts {
    /**
     * Creates a new instance of RichTexts. If cell is provided, adding a {@link RichTextFragment} with
     * text contains line separator will trigger {@link Cell.style}('wrapText', true), which
     * will make MS Excel show the new line. i.e. In MS Excel, Tap "alt+Enter" in a cell, the cell
     * will set wrap text to true automatically. You need to manually set wrapText=true if cell
     * is not provided.
     *
     * @param {Cell|undefined} [cell] - The cell that contains this rich text
     * @param {undefined|null|Object} [node] - The node stored in the shared string
     */
    constructor(cell, node) {
        this._node = [];
        this._cell = cell;
        if (node) {
            for (let i = 0; i < node.length; i++) {
                this._node.push(new RichTextFragment(node[i], null, cell));
            }
        }
    }

    /**
     * Gets which cell this {@link RichTexts} instance belongs to.
     * @return {Cell} The cell this instance belongs to.
     */
    get cell() {
        return this._cell;
    }

    /**
     * Sets which cell this {@link RichTexts} instance belongs to.
     * @see {@link RichTexts}
     * @param {Cell} cell - The cell this instance should belong to.
     */
    set cell(cell) {
        this._cell = cell;
        if (cell && this.text.includes('\n')) {
            cell.style('wrapText', true);
        }
    }

    /**
     * Gets the how many rich text fragment this {@link RichTexts} instance contains
     * @return {number} The number of fragments this {@link RichTexts} instance has.
     */
    get length() {
        return this._node.length;
    }

    /**
     * Gets concatenated text without styles.
     * @return {string} concatenated text
     */
    get text() {
        let text = '';
        for (let i = 0; i < this._node.length; i++) {
            text += this.get(i).value();
        }
        return text;
    }

    /**
     * Gets the ith fragment of this {@link RichTexts} instance.
     * @param {number} index - The index
     * @return {RichTextFragment} A rich text fragment
     */
    get(index) {
        return this._node[index];
    }

    /**
     * Removes a rich text fragment. This instance will be mutated.
     * @param {number} index - the index of the fragment to remove
     * @return {RichTexts} the rich text instance
     */
    remove(index) {
        this._node.splice(index, 1);
        return this;
    }

    /**
     * Adds a rich text fragment to the last or after the given index. This instance will be mutated.
     * @param {string} text - the text
     * @param {{}} [styles] - the styles js object, i.e. {fontSize: 12}
     * @param {number|undefined|null} [index] - the index of the fragment to add
     * @return {RichTexts} the rich text instance
     */
    add(text, styles, index) {
        if (index === undefined || index === null) {
            this._node.push(new RichTextFragment(text, styles, this._cell));
        } else {
            this._node.splice(index, 0, new RichTextFragment(text, styles, this._cell));
        }
        return this;
    }

    /**
     * Clears this rich text
     * @return {RichTexts} the rich text instance
     */
    clear() {
        this._node = [];
        this._cell = undefined;
        return this;
    }

    /**
     * Convert the rich text to an XML object.
     * @returns {{}} The XML form.
     * @ignore
     */
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
