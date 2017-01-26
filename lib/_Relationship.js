"use strict";

const debug = require("./debug")("_Relationship");

// The schema prefix for relationship types
const RELATIONSHIP_SCHEMA_PREFIX = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";

/**
 * A relationship.
 */
class _Relationship {
    /**
     * Creates a new instance of _Relationship.
     * @param {Element} node - The DOM element.
     */
    constructor(node) {
        debug("constructor(_)");
        this._node = node;
    }

    /**
     * Gets the ID.
     * @returns {string} The ID
     */
    id() {
        debug("id()");
        return this._node.getAttribute("Id");
    }

    /**
     * Gets the type of the relationship.
     * @returns {string} The type.
     */
    type() {
        debug("type()");
        return this._node.getAttribute("Type").substr(RELATIONSHIP_SCHEMA_PREFIX.length);
    }
}

module.exports = _Relationship;

/*
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
*/
