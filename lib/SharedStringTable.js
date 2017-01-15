"use strict";

const DOMParser = require('xmldom').DOMParser;
const parser = new DOMParser();

// TODO: Add support for rich text. (Which is why we have to add to the XML doc instead of completely parsing/generating the doc.)

class SharedStringTable {
    constructor(sharedStringsXML) {
        this._stringArray = [];
        this._indexMap = {};

        if (sharedStringsXML) {
            this._xml = sharedStringsXML;
            this._xml.removeAttribute("count");
            this._xml.removeAttribute("uniqueCount");

            for (let i = 0; i < this._xml.childNodes.length; i++) {
                const siNode = this._xml.childNodes[i];
                const text = siNode.firstChild.firstChild.textContent;

                if (text) {
                    const index = this._stringArray.length;
                    this._stringArray.push(text);
                    this._indexMap[text] = index;
                } else {
                    this._stringArray.push(null);
                }
            }
        } else {
            this._xml = parser.parseFromString(`<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></sst>`).documentElement;
        }
    }

    getStringByIndex(index) {
        return this._stringArray[index];
    }

    getIndexForString(string) {
        let index = this._indexMap[string];
        if (index >= 0) return index;

        index = this._stringArray.length;
        this._stringArray.push(string);
        this._indexMap[string] = index;

        const siNode = this._xml.ownerDocument.createElement("si");
        this._xml.appendChild(siNode);
        const tNode = this._xml.ownerDocument.createElement("t");
        siNode.appendChild(tNode);
        const textNode = this._xml.ownerDocument.createTextNode(string);
        tNode.appendChild(textNode);

        return index;
    }

    toString() {
        return this._xml.toString();
    }
}

module.exports = SharedStringTable;
