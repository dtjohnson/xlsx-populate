"use strict";

const _ = require("lodash");
const debug = require("./debug")("Relationships");

const RELATIONSHIP_SCHEMA_PREFIX = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";

/**
 * A relationship collection.
 * @ignore
 */
class Relationships {
    /**
     * Creates a new instance of _Relationships.
     * @param {{}} node - The node.
     */
    constructor(node) {
        debug("constructor(_)");
        this._node = node;
        this._getStartingId();
    }

    /**
     * Add a new relationship.
     * @param {string} type - The type of relationship.
     * @param {string} target - The target of the relationship.
     * @returns {{}} The new relationship.
     */
    add(type, target) {
        debug("add(%o)", arguments);
        const node = {
            name: "Relationship",
            attributes: {
                Id: `rId${this._nextId++}`,
                Type: `${RELATIONSHIP_SCHEMA_PREFIX}${type}`,
                Target: target
            }
        };

        this._node.children.push(node);
        return node;
    }

    /**
     * Find a relationship by type.
     * @param {string} type - The type to search for.
     * @returns {{}|undefined} The matching relationship or undefined if not found.
     */
    findByType(type) {
        debug("findByType(%o)", arguments);
        return _.find(this._node.children, node => node.attributes.Type === `${RELATIONSHIP_SCHEMA_PREFIX}${type}`);
    }

    /**
     * Convert the collection to an object.
     * @returns {{}} The object.
     */
    toObject() {
        debug("toObject(%o)", arguments);
        return this._node;
    }

    /**
     * Get the starting relationship ID to use for new relationships.
     * @private
     * @returns {undefined}
     */
    _getStartingId() {
        this._nextId = 1;
        this._node.children.forEach(node => {
            const id = parseInt(node.attributes.Id.substr(3));
            if (id >= this._nextId) this._nextId = id + 1;
        });
    }
}

module.exports = Relationships;

/*
xl/_rels/workbook.xml.rels

<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
    <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
    <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>
*/

