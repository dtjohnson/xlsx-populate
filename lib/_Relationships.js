"use strict";

const debug = require("./debug")("_Relationships");
const utils = require("./utils");
const _Relationship = require("./_Relationship");
const DOMParser = require('xmldom').DOMParser;
const parser = new DOMParser();

/**
 * A relationship collection.
 */
class _Relationships {
    /**
     * Creates a new instance of _Relationships.
     * @param {string} text - The XML text.
     */
    constructor(text) {
        debug("constructor(_)");
        this._xml = parser.parseFromString(text);

        this._relationships = utils.mapChildElements(this._xml.documentElement, node => {
            return new _Relationship(node);
        });
    }

    /**
     * Add a new relationship.
     * @param {string} type - The type of relationship.
     * @param {string} target - The target of the relationship.
     * @returns {_Relationship} The new relationship.
     */
    add(type, target) {
        debug("add(%o)", arguments);
        const id = Date.now();
        const node = parser.parseFromString(`<Relationship Id="rId${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"/>`);
        this._xml.documentElement.appendChild(node);
        const relationship = new _Relationship(node);
        this._relationships.push(relationship);
        return relationship;
    }

    /**
     * Find a relationship by type.
     * @param {string} type - The type to search for.
     * @returns {_Relationship|undefined} The matching relationship or undefined if not found.
     */
    findByType(type) {
        debug("findByType(%o)", arguments);
        return this._relationships.find(relationship => relationship.type() === type);
    }

    /**
     * Convert the collection to an XML string.
     * @returns {string} The XML string.
     */
    toString() {
        debug("toString(%o)", arguments);
        return this._xml.toString();
    }
}

module.exports = _Relationships;

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

