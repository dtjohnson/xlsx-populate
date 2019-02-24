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
//# sourceMappingURL=ContentTypes.js.map