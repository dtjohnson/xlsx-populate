"use strict";

const DOMParser = require('xmldom').DOMParser;
const parser = new DOMParser();

class _SharedStringsTable {
    constructor(text) {
        this._xml = parser.parseFromString(text);
    }

    toString() {
        return this._xml.toString();
    }
}


module.exports = _SharedStringsTable;

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
