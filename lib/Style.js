"use strict";

const debug = require("./debug")("_Style");

class Style {
    constructor() {
        if (arguments.length === 1) {
            this._styles = arguments[0];
        } else {
            this._id = arguments[0];
            this._xfNode = arguments[1];
            this._fontNode = arguments[2];
        }
    }

    _iterateStyles(methodName, args) {
        const values = [];
        for (let i = 0; i < this._styles.length; i++) {
            values[i] = [];
            for (let j = 0; j < this._styles[i].length; j++) {
                const style = this._styles[i][j];
                const childArgs = [];
                for (let k = 0; k < args.length; k++) {
                    childArgs[k] = Array.isArray(args[k]) ? args[k][i][j] : args[k];
                }

                values[i][j] = style[methodName].apply(style, childArgs);
            }
        }

        return (values[0][0] instanceof Style) ? this : values;
    }

    /**
     * Gets or sets whether the font is bold.
     * @param {boolean} [bold] - The value to set.
     * @returns {boolean|Style} The value if getting or the style if setting.
     */
    bold(bold) {
        debug("bold(bold: %s)", bold);
        if (this._styles) return this._iterateStyles("bold", arguments);
        return this._tagValue("bold", this._fontNode, "b", arguments);
    }

    italic(italic) {
        debug("italic(italic: %s)", italic);
        if (this._styles) return this._iterateStyles("italic", arguments);
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

    _attributeValue(opts) {
        let value = opts.args[0];
        const allAttributeNames = opts.allAttributeNames || [opts.attributeName];

        let node = opts.parentNode.getElementsByTagName(opts.tagName)[0];
        if (opts.args.length === 0) {
            if (!node) return;
            for (let i = 0; i < allAttributeNames.length; i++) {
                const attributeName = allAttributeNames[i];
                value = node.getAttribute(attributeName);
                if (value) {
                    if (opts.fromString) value = opts.fromString(value, attributeName);
                    return value;
                }
            }

            return;
        } else if (opts.args.length === 1) {
            if (!node) {
                node = opts.parentNode.ownerDocument.createElement(opts.tagName);
                opts.parentNode.appendChild(node);
            }

            allAttributeNames.forEach(attributeName => {
                node.removeAttribute(attributeName);
            });

            if (value) node.setAttribute(opts.attributeName, value);
            if (!node.hasAttributes()) opts.parentNode.removeChild(node);

            return this;
        } else {
            throw new Error(`Style.${opts.methodName}: Invalid number of arguments`);
        }
    }

    fontVerticalAlignment(alignment) {
        debug("fontVerticalAlignment(alignment: %s)", alignment);
        return this._attributeValue({
            methodName: "fontVerticalAlignment",
            parentNode: this._fontNode,
            tagName: "vertAlign",
            attributeName: "val",
            args: arguments
        });
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
        debug("fontSize(size: %s)", size);
        return this._attributeValue({
            methodName: "fontSize",
            parentNode: this._fontNode,
            tagName: "sz",
            attributeName: "val",
            args: arguments,
            fromString: val => parseInt(val)
        });
    }

    fontFamily(family) {
        debug("fontFamily(family: %s)", family);
        return this._attributeValue({
            methodName: "fontFamily",
            parentNode: this._fontNode,
            tagName: "name",
            attributeName: "val",
            args: arguments
        });
    }

    fontColor(color) {
        debug("fontColor(color: %s)", color);
        return this._attributeValue({
            methodName: "fontColor",
            parentNode: this._fontNode,
            tagName: "color",
            attributeName: typeof color === "string" ? "rgb" : "indexed",
            args: arguments,
            allAttributeNames: ["rgb", "indexed"],
            fromString: (val, attributeName) => {
                if (attributeName === "indexed") return parseInt(val);
                return val;
            }
        });
    }
}

module.exports = Style;
