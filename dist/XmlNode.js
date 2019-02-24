"use strict";
// TODO: This is a reworked API for XML nodes.
Object.defineProperty(exports, "__esModule", { value: true });
const XML_DECLARATION = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`;
/**
 * Escape a string for use in XML by replacing &, ", ', <, and >.
 * @param {*} value - The value to escape.
 * @param {boolean} [isAttribute] - A flag indicating if this is an attribute.
 * @returns {string} The escaped string.
 * @private
 */
function escapeString(value, isAttribute) {
    value = value.toString()
        .replace(/&/g, "&amp;") // Escape '&' first as the other escapes add them.
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;");
    if (isAttribute) {
        value = value.replace(/"/g, "&quot;");
    }
    return value;
}
class XmlNode {
    constructor(name, attributes) {
        this.name = name;
        if (attributes)
            this.setAttributes(attributes);
    }
    setAttributes(attributes) {
        if (!this.attributes)
            this.attributes = {};
        for (let key in attributes) {
            if (attributes.hasOwnProperty(key)) {
                this.attributes[key] = attributes[key];
            }
        }
    }
    appendChild(child) {
        if (!this.children)
            this.children = [];
        this.children.push(child);
    }
    findChildWithName(name) {
        if (!this.children)
            return;
        for (let i = 0; i < this.children.length; i++) {
            const child = this.children[i];
            if (child instanceof XmlNode && child.name === name)
                return child;
        }
    }
    hasChild(name) {
        return !!this.findChildWithName(name);
    }
    removeChild(child) {
        if (!this.children)
            return false;
        const index = this.children.indexOf(child);
        if (index < 0)
            return false;
        this.children.splice(index, 1);
        return true;
    }
    removeChildWithName(name) {
        const child = this.findChildWithName(name);
        if (!child)
            return false;
        return this.removeChild(child);
    }
    toString(includeDeclaration = false) {
        let str = `<${this.name}`;
        if (this.attributes) {
            for (let key in this.attributes) {
                if (this.attributes.hasOwnProperty(key)) {
                    str += ` ${key}="${escapeString(this.attributes[key], true)}"`;
                }
            }
        }
        if (this.children && this.children.length) {
            str += ">";
            // Recursively add any children.
            this.children.forEach(child => {
                str += (child instanceof XmlNode) ? child : escapeString(child, false);
            });
            // Close the tag.
            str += `</${this.name}>`;
        }
        else {
            // Self-close the tag if no children.
            str += "/>";
        }
        if (includeDeclaration) {
            str = XML_DECLARATION + str;
        }
        return str;
    }
}
exports.XmlNode = XmlNode;
//# sourceMappingURL=XmlNode.js.map