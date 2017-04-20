"use strict";

const _ = require("lodash");

const XML_DECLARATION = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`;

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
        this._i = 0;
        const xml = this._build(node, '');
        if (xml === '') return;
        return XML_DECLARATION + xml;
    }

    /**
     * Build the XML string. (This is the internal recursive method.)
     * @param {{}} node - The node.
     * @param {string} xml - The initial XML doc string.
     * @returns {string} The generated XML element.
     * @private
     */
    _build(node, xml) {
        // For CPU performance, JS engines don't truly concatenate strings; they create a tree of pointers to
        // the various concatenated strings. The downside of this is that it consumes a lot of memory, which
        // will cause problems with large workbooks. So periodically, we grab a character from the xml, which
        // causes the JS engine to flatten the tree into a single string. Do this too often and CPU takes a hit.
        // Too frequently and memory takes a hit. Every 100k nodes seems to be a good balance.
        if (this._i++ % 1000000 === 0) {
            this._c = xml[0];
        }

        // If the node has a toXml method, call it.
        if (node && _.isFunction(node.toXml)) node = node.toXml();

        if (_.isObject(node)) {
            // If the node is an object, then it maps to an element. Check if it has a name.
            if (!node.name) throw new Error(`XML node does not have name: ${JSON.stringify(node)}`);

            // Add the opening tag.
            xml += `<${node.name}`;

            // Add any node attributes
            _.forOwn(node.attributes, (value, name) => {
                xml += ` ${name}="${this._escapeString(value, true)}"`;
            });

            if (_.isEmpty(node.children)) {
                // Self-close the tag if no children.
                xml += "/>";
            } else {
                xml += ">";
                
                // Recursively add any children.
                _.forEach(node.children, child => {
                    // Add the children to the XML.
                    xml = this._build(child, xml);
                });

                // Close the tag.
                xml += `</${node.name}>`;
            }
        } else if (!_.isNil(node)) {
            // It not an object, this should be a text node. Just add it.
            xml += this._escapeString(node);
        }

        // Return the updated XML element.
        return xml;
    }

    /**
     * Escape a string for use in XML by replacing &, ", ', <, and >.
     * @param {*} value - The value to escape.
     * @param {boolean} [isAttribute] - A flag indicating if this is an attribute.
     * @returns {string} The escaped string.
     * @private
     */
    _escapeString(value, isAttribute) {
        if (_.isNil(value)) return value;
        value = value.toString()
            .replace(/&/g, "&amp;") // Escape '&' first as the other escapes add them.
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;");

        if (isAttribute) {
            value = value.replace(/"/g, "&quot;");
        }

        return value;
    }
}

module.exports = XmlBuilder;
