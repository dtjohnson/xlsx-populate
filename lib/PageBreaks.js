"use strict";

/**
 * PageBreaks
 */
class PageBreaks {
    constructor(node) {
        this._node = node;
    }

    /**
     * add page-breaks by row/column id
     * @param {number} id - row/column id (rowNumber/colNumber)
     * @return {PageBreaks} the page-breaks
     */
    add(id) {
        this._node.children.push({
            name: "brk",
            children: [],
            attributes: {
                id,
                max: 16383,
                man: 1
            }
        });
        this._node.attributes.count++;
        this._node.attributes.manualBreakCount++;

        return this;
    }

    /**
     * remove page-breaks by index
     * @param {number} index - index of list
     * @return {PageBreaks} the page-breaks
     */
    remove(index) {
        const brk = this._node.children[index];
        if (brk) {
            this._node.children.splice(index, 1);
            this._node.attributes.count--;
            if (brk.man) {
                this._node.attributes.manualBreakCount--;
            }
        }

        return this;
    }

    /**
     * get count of the page-breaks
     * @return {number} the page-breaks' count
     */
    get count() {
        return this._node.attributes.count;
    }

    /**
     * get list of page-breaks
     * @return {Array} list of the page-breaks
     */
    get list() {
        return this._node.children.map(brk => ({
            id: brk.id,
            isManual: !!brk.man
        }));
    }
}

module.exports = PageBreaks;
