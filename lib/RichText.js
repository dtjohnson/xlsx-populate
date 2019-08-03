"use strict";

const _ = require("lodash");
const RichTextFragment = require("./RichTextFragment");

/**
 * A RichText class that contains many {@link RichTextFragment}.
 */
class RichText {
    /**
     * Creates a new instance of RichText. If you get the instance by calling `Cell.value()`,
     * adding a text contains line separator will trigger {@link Cell.style}('wrapText', true), which
     * will make MS Excel show the new line. i.e. In MS Excel, Tap "alt+Enter" in a cell, the cell
     * will set wrap text to true automatically.
     *
     * @param {undefined|null|Object} [node] - The node stored in the shared string
     */
    constructor(node) {
        this._node = [];
        this._cell = null;
        this._remainingNodes = [];
        if (node) {
            for (let i = 0; i < node.length; i++) {
                const fragment = node[i];
                if (fragment.name === 'r') {
                    this._node.push(new RichTextFragment(fragment, null, this));
                } else {
                    // special node, e.g. rPh, phoneticPr in Japanese language.
                    this._remainingNodes.push(fragment);
                }
            }
        }
    }

    /**
     * Gets which cell this {@link RichText} instance belongs to.
     * @return {Cell|undefined} The cell this instance belongs to.
     */
    get cell() {
        return this._cell;
    }

    /**
     * Gets the how many rich text fragment this {@link RichText} instance contains
     * @return {number} The number of fragments this {@link RichText} instance has.
     */
    get length() {
        return this._node.length;
    }

    /**
     * Gets concatenated text without styles.
     * @return {string} concatenated text
     */
    text() {
        let text = '';
        for (let i = 0; i < this._node.length; i++) {
            text += this.get(i).value();
        }
        return text;
    }

    /**
     * Gets the instance with cell reference defined.
     * @param {Cell} cell - Cell reference.
     * @return {RichText} The instance with cell reference defined.
     */
    getInstanceWithCellRef(cell) {
        this._cell = cell;
        return this;
    }

    /**
     * Returns a deep copy of this instance.
     * If cell reference is provided, it checks line separators and calls
     * `cell.style('wrapText', true)` when needed.
     * @param {Cell|undefined} [cell] - The cell reference.
     * @return {RichText} A deep copied instance
     */
    copy(cell) {
        const newRichText = new RichText(_.cloneDeep(this.toXml()));
        if (cell && _.includes(this.text(), '\n')) {
            cell.style('wrapText', true);
        }
        return newRichText;
    }

    /**
     * Gets the ith fragment of this {@link RichText} instance.
     * @param {number} index - The index
     * @return {RichTextFragment} A rich text fragment
     */
    get(index) {
        return this._node[index];
    }

    /**
     * Removes a rich text fragment. This instance will be mutated.
     * @param {number} index - the index of the fragment to remove
     * @return {RichText} the rich text instance
     */
    remove(index) {
        this._node.splice(index, 1);
        this.removeUnsupportedNodes();
        return this;
    }

    /**
     * Adds a rich text fragment to the last or after the given index. This instance will be mutated.
     * @param {string} text - the text
     * @param {{}} [styles] - the styles js object, i.e. {fontSize: 12}
     * @param {number|undefined|null} [index] - the index of the fragment to add
     * @return {RichText} the rich text instance
     */
    add(text, styles, index) {
        if (index === undefined || index === null) {
            this._node.push(new RichTextFragment(text, styles, this));
        } else {
            this._node.splice(index, 0, new RichTextFragment(text, styles, this));
        }
        this.removeUnsupportedNodes();
        return this;
    }

    /**
     * Clears this rich text
     * @return {RichText} the rich text instance
     */
    clear() {
        this._node = [];
        this._remainingNodes = [];
        this._cell = undefined;
        return this;
    }

    /**
     * Remove all unsupported nodes (phoneticPr, rPh for Japanese language).
     * @return {undefined}
     */
    removeUnsupportedNodes() {
        this._remainingNodes = [];
    }

    /**
     * Convert the rich text to an XML object.
     * @returns {Array.<{}>} The XML form.
     * @ignore
     */
    toXml() {
        const node = [];
        for (let i = 0; i < this._node.length; i++) {
            node.push(this._node[i].toXml());
        }
        return node.concat(this._remainingNodes);
    }
}

// IE doesn't support function names so explicitly set it.
if (!RichText.name) RichText.name = "RichText";

module.exports = RichText;
