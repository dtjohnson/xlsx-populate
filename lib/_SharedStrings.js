"use strict";

const debug = require("./debug")('_SharedStrings');
const jq = require("./jq");

/**
 * The shared strings table.
 */
class _SharedStrings {
    /**
     * Constructs a new instance of _SharedStrings.
     * @param {{}} node - The node.
     */
    constructor(node) {
        debug("constructor(_)");
        this._stringArray = [];
        this._indexMap = {};

        this._initNode(node);
        this._cacheExistingSharedStrings();
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
        this._siNode.push({ t: [string] });

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
     * Convert the collection to an object.
     * @returns {{}} The object.
     */
    toObject() {
        debug("toObject(%o)", arguments);
        return this._node;
    }

    /**
     * Store any existing values in the caches.
     * @private
     * @returns {undefined}
     */
    _cacheExistingSharedStrings() {
        debug("_cacheExistingSharedStrings(%o)", arguments);
        this._siNode.forEach((si, i) => {
            if (si.t) {
                const string = si.t[0];
                this._stringArray.push(string);
                this._indexMap[string] = i;
            } else {
                // TODO: Properly support rich text nodes in the future. For now just store the object as a placeholder.
                this._stringArray.push(si);
            }
        });
    }

    /**
     * Initialize the node.
     * @param {{}} [node] - The shared strings node.
     * @private
     * @returns {undefined}
     */
    _initNode(node) {
        debug("_initNode(_)");
        if (!node) node = {
            sst: {
                $: {
                    xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                },
                si: []
            }
        };

        this._node = node;
        this._siNode = this._node.sst.si;

        jq.set(this._node.sst, {
            "$.count": null,
            "$.unique": null
        });
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
