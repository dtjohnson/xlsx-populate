"use strict";

// TODO: Tests
// TODO: JSDoc
// TODO: Debugs

const RELATIONSHIP_SCHEMA_PREFIX = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";

class _Relationship {
    constructor(node) {
        this._node = node;
    }

    id() {
        return this._node.getAttribute("Id");
    }

    type() {
        return this._node.getAttribute("Type").substr(RELATIONSHIP_SCHEMA_PREFIX.length);
    }
}

module.exports = _Relationship;

/*
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
*/
