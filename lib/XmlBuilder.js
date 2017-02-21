"use strict";

const builder = require('xmlbuilder');
const _ = require("lodash");

/**
 * XML document builder.
 * @private
 */
class XmlBuilder {
    /**
     * Build an XML string from the JSON object.
     * @param {{}} node - The node.
     * @returns {string} The XML text.
     */
    build(node) {
        const xml = this._build(node);
        return xml.end({ pretty: true });
    }

    /**
     * Build the XML string. (This is the internal recursive method.)
     * @param {{}} node - The node.
     * @param {XmlElement} xml - The current XML element.
     * @returns {XmlElement} - The generated XML element.
     * @private
     */
    _build(node, xml) {
        if (_.isObject(node)) {
            // If the node is an object, then it maps to an element. Check if it has a name.
            if (!node.name) throw new Error("XML node does not have name");

            // If the XML element is already set, we want to add a child element. Otherwise, this should be the root element.
            if (xml) {
                xml = xml.element(node.name);
            } else {
                xml = builder.create(node.name, { encoding: "UTF-8", standalone: true });
            }

            // Add any node attributes.
            _.forOwn(node.attributes, (value, name) => {
                xml.attribute(name, value);
            });

            // Recursively add any children.
            _.forEach(node.children, child => this._build(child, xml));
        } else {
            // It not an object, this should be a text node.
            xml.text(node);
        }

        // Return the updated XML element.
        return xml;
    }
}

module.exports = XmlBuilder;
