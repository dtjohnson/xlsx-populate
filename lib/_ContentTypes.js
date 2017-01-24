"use strict";

// TODO: Tests
// TODO: JSDoc
// TODO: Debugs

const _ContentType = require("./_ContentType");
const DOMParser = require('xmldom').DOMParser;
const parser = new DOMParser();

class _ContentTypes {
    constructor(text) {
        this._xml = parser.parseFromString(text);

        this._contentTypes = [];
        for (let i = 0; i < this._xml.documentElement.childNodes.length; i++) {
            const node = this._xml.documentElement.childNodes[i];
            this._contentTypes.push(new _ContentType(node));
        }
    }

    findByPartName(partName) {
        return this._contentTypes.find(contentType => contentType.partName() === partName);
    }

    add(partName, contentType) {
        const node = parser.parseFromString(`<Override PartName="${partName}" ContentType="${contentType}"/>`)
        this._xml.documentElement.appendChild(node);
        this._contentTypes.push(new _ContentType(node));
    }

    toString() {
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
