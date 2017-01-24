"use strict";

// TODO: Tests
// TODO: JSDoc
// TODO: Switch to xq.query/update
// TODO: Debugs

const DOMParser = require('xmldom').DOMParser;
const parser = new DOMParser();

class _SharedStrings {
    constructor(text) {
        this._stringArray = [];
        this._indexMap = {};

        if (!text) text = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></sst>`;
        this._xml = parser.parseFromString(text);

        this._tableNode = this._xml.documentElement;
        this._tableNode.removeAttribute("count");
        this._tableNode.removeAttribute("uniqueCount");

        for (let i = 0; i < this._tableNode.childNodes.length; i++) {
            const siNode = this._tableNode.childNodes[i];
            const text = siNode.firstChild.firstChild.textContent;

            if (text) {
                const index = this._stringArray.length;
                this._stringArray.push(text);
                this._indexMap[text] = index;
            } else {
                this._stringArray.push(null);
            }
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

        const siNode = this._xml.createElement("si");
        this._tableNode.appendChild(siNode);
        const tNode = this._xml.createElement("t");
        siNode.appendChild(tNode);
        const textNode = this._xml.createTextNode(string);
        tNode.appendChild(textNode);

        return index;
    }

    toString() {
        return this._xml.toString();
    }
}


module.exports = _SharedStrings;

/*
xl/sharedStrings.xml

<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="13" uniqueCount="4">
	<si>
		<t>Foo</t>
	</si>
	<si>
		<t>Bar</t>
	</si>
	<si>
		<t>Goo</t>
	</si>
	<si>
		<r>
			<t>s</t>
		</r><r>
			<rPr>
				<b/>
				<sz val="11"/>
				<color theme="1"/>
				<rFont val="Calibri"/>
				<family val="2"/>
				<scheme val="minor"/>
			</rPr><t>d;</t>
		</r><r>
			<rPr>
				<sz val="11"/>
				<color theme="1"/>
				<rFont val="Calibri"/>
				<family val="2"/>
				<scheme val="minor"/>
			</rPr><t>lfk;l</t>
		</r>
	</si>
</sst>
*/
