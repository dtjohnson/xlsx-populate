"use strict";

const debug = require("./debug")('_SharedStrings');
const DOMParser = require('xmldom').DOMParser;
const parser = new DOMParser();
const utils = require("./utils");
const xq = require("./xq");

/**
 * The shared strings table.
 */
class _SharedStrings {
    /**
     * Constructs a new instance of _SharedStrings.
     * @param {string} text - The XML text from xl/sharedStrings.xml
     */
    constructor(text) {
        debug("constructor(_)");

        this._stringArray = [];
        this._indexMap = {};

        // The shared string table is not mandatory. If it doesn't exist, create it.
        if (!text) text = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></sst>`;
        this._xml = parser.parseFromString(text);

        this._tableNode = this._xml.documentElement;

        // Remove the counts as they don't seem necessary.
        xq.update(this._tableNode, {
            "@count": null,
            "@uniqueCount": null
        });

        // Query for the existing values.
        const result = xq.query(this._tableNode, {
            si: {
                $multi: true,
                t: {
                    $optional: true,
                    "#text": String
                }
            }
        });

        // Store any existing values in the caches.
        if (result) result.si.forEach((si, i) => {
            if (si.t) {
                const string = si.t['#text'];
                this._stringArray.push(string);
                this._indexMap[string] = i;
            } else {
                // TODO: Support rich text nodes in the future. For now just store a null as a placeholder.
                this._stringArray.push(null);
            }
        });
    }

    /**
     * Gets the index for a string
     * @param {string} string - The string
     * @returns {number} The index
     */
    getIndexForString(string) {
        debug("getIndexForString(%o)", arguments);

        // If the string is found in the cache, return the index.
        let index = this._indexMap[string];
        if (index >= 0) return index;

        // Otherwise, add it to the caches.
        index = this._stringArray.length;
        this._stringArray.push(string);
        this._indexMap[string] = index;

        // Append a new si node.
        xq.update(this._tableNode, {
            si: {
                $append: true,
                t: {
                    "#text": string
                }
            }
        });

        return index;
    }

    /**
     * Get the string for a given index
     * @param {number} index - The index
     * @returns {string} The string
     */
    getStringByIndex(index) {
        debug("getStringByIndex(%o)", arguments);
        return this._stringArray[index];
    }

    /**
     * Converts to an XML string.
     * @returns {string} The XML.
     */
    toString() {
        debug("toString(%o)", arguments);
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
