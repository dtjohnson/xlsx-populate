"use strict";

const debug = require("debug")("xlsx-populate:_StyleSheet");

class Style {
    constructor(id, xfNode, fontNode) {
        this._id = id;
        this._xfNode = xfNode;
        this._fontNode = fontNode;
    }

    bold(bold) {
        return this._tagValue("bold", this._fontNode, "b", arguments);
    }

    italic(italic) {
        return this._tagValue("italic", this._fontNode, "i", arguments);
    }

    underline(underline) {
        return this._tagValue("underline", this._fontNode, "u", arguments);
    }

    strikethrough(strikethrough) {
        return this._tagValue("strikethrough", this._fontNode, "strike", arguments);
    }

    _tagValue(methodName, parentNode, tagName, args) {
        const value = args[0];
        let node = parentNode.getElementsByTagName(tagName)[0];
        if (args.length === 0) {
            return Boolean(node);
        } else if (args.length === 1) {
            if (value && !node) {
                node = parentNode.ownerDocument.createElement(tagName);
                parentNode.appendChild(node);
            } else if (!value && node) {
                parentNode.removeChild(node);
            }

            return this;
        } else {
            throw new Error(`Style.${methodName}: Invalid number of arguments`);
        }
    }

    _attributeValue(methodName, parentNode, tagName, attributeName, args) {
        const value = args[0];
        let node = parentNode.getElementsByTagName(tagName)[0];
        if (args.length === 0) {
            return node && node.getAttribute(attributeName);
        } else if (args.length === 1) {
            if (!node) {
                node = parentNode.ownerDocument.createElement(tagName);
                parentNode.appendChild(node);
            }

            if (value) {
                node.setAttribute(attributeName, value);
            } else {
                node.removeAttribute(attributeName);
                if (!node.hasAttributes()) parentNode.removeChild(node);
            }

            return this;
        } else {
            throw new Error(`Style.${methodName}: Invalid number of arguments`);
        }
    }

    fontVerticalAlignment(alignment) {
        debug("fontVerticalAlignment(alignment: %s)", alignment);
        return this._attributeValue("fontVerticalAlignment", this._fontNode, "vertAlign", "val", arguments);
    }

    superscript(superscript) {
        debug("superscript(superscript: %s)", superscript);

        if (arguments.length === 0) {
            return this.fontVerticalAlignment() === "superscript";
        } else if (arguments.length === 1) {
            return this.fontVerticalAlignment(superscript && "superscript");
        } else {
            throw new Error("Style.superscript: Invalid number of arguments");
        }
    }

    subscript(subscript) {
        debug("subscript(subscript: %s)", subscript);

        if (arguments.length === 0) {
            return this.fontVerticalAlignment() === "subscript";
        } else if (arguments.length === 1) {
            return this.fontVerticalAlignment(subscript && "subscript");
        } else {
            throw new Error("Style.subscript: Invalid number of arguments");
        }
    }

    fontSize(size) {
        debug("fontVerticalAlignment(size: %s)", size);
        return this._attributeValue("fontSize", this._fontNode, "sz", "val", arguments);
    }

    fontFamily(family) {
        debug("fontVerticalAlignment(family: %s)", family);
        return this._attributeValue("fontFamily", this._fontNode, "name", "val", arguments);
    }
}

module.exports = Style;
