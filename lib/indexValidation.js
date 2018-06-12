"use strict";

module.exports = {
    validateIndex(index, kind) {
        if (index < 1) {
            throw new RangeError("Invaid " + kind + " index " + index + ". Remember that spreadsheets use 1 indexing.");
        }
    }
};
