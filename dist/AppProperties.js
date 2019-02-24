"use strict";
const _ = require("lodash");
const xmlq = require("./xmlq");
const ArgHandler = require("./ArgHandler");
/**
 * App properties
 * @ignore
 */
class AppProperties {
    /**
     * Creates a new instance of AppProperties
     * @param {{}} node - The node.
     */
    constructor(node) {
        this._node = node;
    }
    isSecure(value) {
        return new ArgHandler("Range.formula")
            .case(() => {
            const docSecurityNode = xmlq.findChild(this._node, "DocSecurity");
            if (!docSecurityNode)
                return false;
            return docSecurityNode.children[0] === 1;
        })
            .case('boolean', value => {
            const docSecurityNode = xmlq.appendChildIfNotFound(this._node, "DocSecurity");
            docSecurityNode.children = [value ? 1 : 0];
            return this;
        })
            .handle(arguments);
    }
    /**
     * Convert the collection to an XML object.
     * @returns {{}} The XML.
     */
    toXml() {
        return this._node;
    }
}
module.exports = AppProperties;
//# sourceMappingURL=AppProperties.js.map