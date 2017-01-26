"use strict";

const debug = require("./debug")("_ContentType");

/**
 * A content type.
 */
class _ContentType {
    /**
     * Creates an instance of _ContentType
     * @param {Element} node - The node.
     */
    constructor(node) {
        debug("constructor(_)");
        this._node = node;
    }

    /**
     * Get the content type.
     * @returns {string} The content type.
     */
    contentType() {
        debug("contentType()");
        return this._node.getAttribute("ContentType");
    }

    /**
     * Get the name of the part.
     * @returns {string} The part name.
     */
    partName() {
        debug("partName()");
        return this._node.getAttribute("PartName");
    }
}

module.exports = _ContentType;

/*
 <Default Extension="xml" ContentType="application/xml"/>
 OR
 <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
 */
