"use strict";

const _ = require("lodash");

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
        this._init(node);
        this._getStartingId();
    }

    /**
     * Add a new relationship.
     * @param {string} type - The type of relationship.
     * @param {string} target - The target of the relationship.
     * @param {string} [targetMode] - The target mode of the relationship.
     * @returns {{}} The new relationship.
     */
    add(type, target, targetMode) {
        const node = {
            name: "Relationship",
            attributes: {
                Id: `rId${this._nextId++}`,
                Type: `${RELATIONSHIP_SCHEMA_PREFIX}${type}`,
                Target: target
            }
        };

        if (targetMode) {
            node.attributes.TargetMode = targetMode;
        }

        this._node.children.push(node);
        return node;
    }

    /**
     * Find a relationship by ID.
     * @param {string} id - The relationship ID.
     * @returns {{}|undefined} The matching relationship or undefined if not found.
     */
    findById(id) {
        return _.find(this._node.children, node => node.attributes.Id === id);
    }

    /**
     * Find a relationship by type.
     * @param {string} type - The type to search for.
     * @returns {{}|undefined} The matching relationship or undefined if not found.
     */
    findByType(type) {
        return _.find(this._node.children, node => node.attributes.Type === `${RELATIONSHIP_SCHEMA_PREFIX}${type}`);
    }

    /**
     * Convert the collection to an XML object.
     * @returns {{}|undefined} The XML or undefined if empty.
     */
    toXml() {
        if (!this._node.children.length) return;
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

    /**
     * Initialize the node.
     * @param {{}} [node] - The relationships node.
     * @private
     * @returns {undefined}
     */
    _init(node) {
        if (!node) node = {
            name: "Relationships",
            attributes: {
                xmlns: "http://schemas.openxmlformats.org/package/2006/relationships"
            },
            children: []
        };

        this._node = node;
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

