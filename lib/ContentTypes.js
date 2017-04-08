"use strict";

const _ = require("lodash");

/**
 * A content type collection.
 * @ignore
 */
class ContentTypes {
    /**
     * Creates a new instance of ContentTypes
     * @param {{}} node - The node.
     */
    constructor(node) {
        this._node = node;
    }

    /**
     * Add a new content type.
     * @param {string} partName - The part name.
     * @param {string} contentType - The content type.
     * @returns {{}} The new content type.
     */
    add(partName, contentType) {
        const node = {
            name: "Override",
            attributes: {
                PartName: partName,
                ContentType: contentType
            }
        };

        this._node.children.push(node);
        return node;
    }

    /**
     * Find a content type by part name.
     * @param {string} partName - The part name.
     * @returns {{}|undefined} The matching content type or undefined if not found.
     */
    findByPartName(partName) {
        return _.find(this._node.children, node => node.attributes.PartName === partName);
    }

    /**
     * Convert the collection to an XML object.
     * @returns {{}} The XML.
     */
    toXml() {
        return this._node;
    }
}

module.exports = ContentTypes;

/*
[Content_Types].xml

<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"/>
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
    <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
    <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
    <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
    <Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>
    <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
    <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
*/
