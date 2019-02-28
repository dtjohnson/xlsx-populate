/**
 * @module xlsx-populate
 */

import { INode, NodeChild } from './XmlParser';

/**
 * The shared strings table.
 * @ignore
 */
export class SharedStrings {
    private _stringArray: (string|number|NodeChild[])[] = [];
    private _indexMap: { [str: string]: number } = {};
    private _node: INode;

    /**
     * Constructs a new instance of SharedStrings.
     * @param node - The node.
     */
    public constructor(node?: INode) {
        this._node = node || {
            name: 'sst',
            attributes: {
                xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
            },
        };

        if (this._node.attributes) {
            delete this._node.attributes.count;
            delete this._node.attributes.uniqueCount;
        }

        this._cacheExistingSharedStrings();
    }

    /**
     * Gets the index for a string
     * @param str - The string or rich text array.
     * @returns The index
     * // TODO Handle Array.<{}>
     */
    public getIndexForString(str: string|number|INode[]): number {
        // If the string is found in the cache, return the index.
        const key = Array.isArray(str) ? JSON.stringify(str) : str;
        let index = this._indexMap[key];
        if (index >= 0) return index;

        // Otherwise, add it to the caches.
        index = this._stringArray.length;
        this._stringArray.push(str);
        this._indexMap[key] = index;

        // Append a new si node.
        if (!this._node.children) this._node.children = [];
        this._node.children.push({
            name: 'si',
            children: Array.isArray(str) ? str : [
                {
                    name: 't',
                    attributes: { 'xml:space': 'preserve' },
                    children: [ str ],
                },
            ],
        });

        return index;
    }

    /**
     * Get the string for a given index
     * @param index - The index
     * @returns The string
     */
    public getStringByIndex(index: number): string|number|NodeChild[] {
        return this._stringArray[index];
    }

    /**
     * Convert the collection to an XML object.
     * @returns The XML object.
     */
    public toXml(): INode {
        return this._node;
    }

    /**
     * Store any existing values in the caches.
     */
    private _cacheExistingSharedStrings(): void {
        if (this._node.children) {
            this._node.children.forEach((child, i) => {
                // TODO: Need helper methods to make this less vebose
                if (typeof child === 'string'
                    || typeof child === 'number'
                    || !child.children
                    || !child.children.length) return;

                const content = child.children[0];
                if (typeof content === 'string'
                    || typeof content === 'number'
                    || !content.children
                    || !content.children.length) return;

                const str = content.children[0];
                if (content.name === 't' && content.children.length === 1
                    && (typeof str === 'string' || typeof str === 'number')) {
                    this._stringArray.push(str);
                    this._indexMap[str] = i;
                } else {
                    // TODO: Properly support rich text nodes in the future. For now just store the object as a placeholder.
                    this._stringArray.push(child.children);
                    this._indexMap[JSON.stringify(child.children)] = i;
                }
            });
        }
    }
}

// tslint:disable
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
