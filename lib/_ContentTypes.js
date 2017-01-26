"use strict";

const debug = require("./debug")("_ContentTypes");
const utils = require("./utils");
const _ContentType = require("./_ContentType");
const DOMParser = require('xmldom').DOMParser;
const parser = new DOMParser();

/**
 * A content type collection.
 */
class _ContentTypes {
    /**
     * Creates a new instance of _ContentTypes
     * @param {string} text - The XML text.
     */
    constructor(text) {
        debug("constructor(_)");
        this._xml = parser.parseFromString(text);

        this._contentTypes = utils.mapChildElements(this._xml.documentElement, (node, i) => {
            return new _ContentType(node);
        });
    }

    /**
     * Add a new content type.
     * @param {string} partName - The part name.
     * @param {string} contentType - The content type.
     * @returns {_ContentType} The new content type.
     */
    add(partName, contentType) {
        debug("add(%o)", arguments);
        const node = parser.parseFromString(`<Override PartName="${partName}" ContentType="${contentType}"/>`);
        this._xml.documentElement.appendChild(node);
        const contentTypeObj = new _ContentType(node);
        this._contentTypes.push(contentTypeObj);
        return contentTypeObj;
    }

    /**
     * Find a content type by part name.
     * @param {string} partName - The part name.
     * @returns {_ContentType|null} The matching content type or null if not found.
     */
    findByPartName(partName) {
        debug("findByPartName(%o)", arguments);
        return this._contentTypes.find(contentType => contentType.partName() === partName);
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

module.exports = _ContentTypes;

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
