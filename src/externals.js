"use strict";

const JSZip = require("jszip");

/**
 * External modules.
 * @private
 */
module.exports = {
    /**
     * The Promise library.
     * @type {Promise}
     */
    get Promise() {
        return JSZip.external.Promise;
    },

    set Promise(Promise) {
        JSZip.external.Promise = Promise;
    }
};
