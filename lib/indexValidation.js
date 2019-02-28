"use strict";

/**
 * Index validation.
 * @private
 */
module.exports = {
    /**
     * Validate the given index against the 1 indexing system used for
     * spreadsheets. If the index is below 1, then a RangeError is thrown with
     * a helpful error message.
     * @param {number} index - The index to validate.
     * @param {string} kind - The type of index (row, column, etc.).
     */
    validateIndex(index, kind) {
        if (index < 1) {
            throw new RangeError("Invaid " + kind + " index " + index + ". Remember that spreadsheets use 1 indexing.");
        }
    }
};
