"use strict";

const _Relationship = require("./_Relationship");
const DOMParser = require('xmldom').DOMParser;
const parser = new DOMParser();

class _Relationships {
    constructor(text) {
        this._xml = parser.parseFromString(text);

        this._relationships = [];
        for (let i = 0; i < this._xml.documentElement.childNodes.length; i++) {
            const node = this._xml.documentElement.childNodes[i];
            this._relationships.push(new _Relationship(node));
        }
    }

    findByType(type) {
        return this._relationships.find(relationship => relationship.type() === type);
    }

    add(type, target) {
        const id = Date.now();
        const node = parser.parseFromString(`<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"/>`)
        this._xml.documentElement.appendChild(node);
        this._relationships.push(new _Relationships(node));
    }

    remove(relationship) {
        this._relationships.splice(this._relationships.indexOf(relationship), 1);
        this._xml.documentElement.removeChild(relationship._node);
    }

    toString() {
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

