"use strict";

const debug = require("./debug")("_Style");
const xq = require("./xq");

class Style {
    constructor() {
        if (arguments.length === 1) {
            this._styles = arguments[0];
        } else {
            this._id = arguments[0];
            this._xfNode = arguments[1];
            this._fontNode = arguments[2];
            this._borderNode = arguments[3];
        }
    }

    style() {
        debug("style(%o)", arguments);

        if (arguments.length === 1 && typeof arguments[0] === "string") {
            const styleName = arguments[0];
            return this[`__${styleName}`]();
        } else if (arguments.length === 2 && typeof arguments[0] === "string") {
            const styleName = arguments[0];
            const value = arguments[1];
            this[`__${styleName}`](value);
            return this;
        } else if (arguments.length === 1 && Array.isArray(arguments[0])) {
            const result = {};
            arguments[0].forEach(style => {
                result[style] = this.style(style);
            });

            return result;
        } else if (arguments.length === 1 && arguments[0] && arguments[0].constructor === Object) {
            const styles = arguments[0];
            for (const style in styles) {
                if (!styles.hasOwnProperty(style)) continue;
                const value = styles[style];
                this.style(style, value);
            }

            return this;
        }
    }

    __bold(value) {
        debug("__bold(%o)", arguments);
        if (arguments.length === 0) return !!xq.query(this._fontNode, { b: {} });
        xq.update(this._fontNode, { b: value ? {} : null });
    }

    _iterateStyles(methodName, args) {
        debug("_iterateStyles(%s, ...)", methodName)
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

        return values[0][0] instanceof Style ? this : values;
    }

    /**
     * Gets or sets whether the font is bold.
     * @param {boolean} [bold] - The value to set.
     * @returns {boolean|Style} The value if getting or the style if setting.
     */
    bold(bold) {
        debug("bold(%o)", arguments);
        if (this._styles) return this._iterateStyles("bold", arguments);
        return this._tagValue({
            methodName: "bold",
            parentNode: this._fontNode,
            tagName: "b",
            args: arguments
        });
    }

    italic(italic) {
        debug("italic(%o)", arguments);
        if (this._styles) return this._iterateStyles("italic", arguments);
        return this._tagValue({
            methodName: "italic",
            parentNode: this._fontNode,
            tagName: "i",
            args: arguments
        });
    }

    underline(underline) {
        debug("underline(%o)", arguments);
        if (this._styles) return this._iterateStyles("underline", arguments);
        return this._tagValue({
            methodName: "underline",
            parentNode: this._fontNode,
            tagName: "u",
            args: arguments
        });
    }

    strikethrough(strikethrough) {
        debug("strikethrough(%o)", arguments);
        if (this._styles) return this._iterateStyles("strikethrough", arguments);
        return this._tagValue({
            methodName: "strikethrough",
            parentNode: this._fontNode,
            tagName: "strike",
            args: arguments
        });
    }

    _tagValue(opts) {
        debug("_tagValue(...)");
        const value = opts.args[0];
        let node = opts.parentNode.getElementsByTagName(opts.tagName)[0];
        if (opts.args.length === 0) {
            return Boolean(node);
        } else if (opts.args.length === 1) {
            if (value && !node) {
                node = opts.parentNode.ownerDocument.createElement(opts.tagName);
                opts.parentNode.appendChild(node);
            } else if (!value && node) {
                opts.parentNode.removeChild(node);
            }

            return this;
        } else {
            throw new Error(`Style.${opts.methodName}: Invalid number of arguments`);
        }
    }

    _attributeValue(opts) {
        debug("_attributeValue(...)");
        let value = opts.args[0];
        const allAttributeNames = opts.allAttributeNames || [opts.attributeName];

        let node = opts.parentNode.getElementsByTagName(opts.tagName)[0];
        if (opts.args.length === 0) {
            if (!node) return;
            for (let i = 0; i < allAttributeNames.length; i++) {
                const attributeName = allAttributeNames[i];
                value = node.getAttribute(attributeName);
                if (value) {
                    if (opts.fromStringConverter) value = opts.fromStringConverter(value, attributeName);
                    return value;
                }
            }
        } else if (opts.args.length === 1) {
            if (!node) {
                node = opts.parentNode.ownerDocument.createElement(opts.tagName);
                opts.parentNode.appendChild(node);
            }

            allAttributeNames.forEach(attributeName => {
                node.removeAttribute(attributeName);
            });

            if (value) node.setAttribute(opts.attributeName, opts.toStringConverter ? opts.toStringConverter(value) : value);
            if (!node.hasAttributes()) opts.parentNode.removeChild(node);

            return this;
        } else {
            throw new Error(`Style.${opts.methodName}: Invalid number of arguments`);
        }
    }

    fontVerticalAlignment(alignment) {
        debug("fontVerticalAlignment(%o)", arguments);
        if (this._styles) return this._iterateStyles("fontVerticalAlignment", arguments);
        return this._attributeValue({
            methodName: "fontVerticalAlignment",
            parentNode: this._fontNode,
            tagName: "vertAlign",
            attributeName: "val",
            args: arguments
        });
    }

    _shortcutValue(opts) {
        debug("_shortcutValue(...)");
        if (opts.args.length === 0) {
            return this[opts.upstreamMethodName]() === opts.value;
        } else if (opts.args.length === 1) {
            return this[opts.upstreamMethodName](opts.args[0] && opts.value);
        } else {
            throw new Error(`Style.${opts.methodName}: Invalid number of arguments`);
        }
    }

    superscript(superscript) {
        debug("superscript(%o)", arguments);
        if (this._styles) return this._iterateStyles("superscript", arguments);
        return this._shortcutValue({
            methodName: "superscript",
            upstreamMethodName: "fontVerticalAlignment",
            value: "superscript",
            args: arguments
        });
    }

    subscript(subscript) {
        debug("subscript(%o)", arguments);
        if (this._styles) return this._iterateStyles("subscript", arguments);
        return this._shortcutValue({
            methodName: "subscript",
            upstreamMethodName: "fontVerticalAlignment",
            value: "subscript",
            args: arguments
        });
    }

    fontSize(size) {
        debug("fontSize(%o)", arguments);
        if (this._styles) return this._iterateStyles("fontSize", arguments);
        return this._attributeValue({
            methodName: "fontSize",
            parentNode: this._fontNode,
            tagName: "sz",
            attributeName: "val",
            args: arguments,
            fromStringConverter: val => parseInt(val)
        });
    }

    fontFamily(family) {
        debug("fontFamily(%o)", arguments);
        if (this._styles) return this._iterateStyles("fontFamily", arguments);
        return this._attributeValue({
            methodName: "fontFamily",
            parentNode: this._fontNode,
            tagName: "name",
            attributeName: "val",
            args: arguments
        });
    }

    fontColor(color) {
        debug("fontColor(%o)", arguments);
        if (this._styles) return this._iterateStyles("fontColor", arguments);
        return this._attributeValue({
            methodName: "fontColor",
            parentNode: this._fontNode,
            tagName: "color",
            attributeName: typeof color === "string" ? "rgb" : "indexed",
            args: arguments,
            allAttributeNames: ["rgb", "indexed"],
            fromStringConverter: (val, attributeName) => {
                if (attributeName === "indexed") return parseInt(val);
                return val;
            }
        });
    }

    horizontalAlignment(alignment) {
        debug("horizontalAlignment(%o)", arguments);
        if (this._styles) return this._iterateStyles("horizontalAlignment", arguments);
        return this._attributeValue({
            methodName: "horizontalAlignment",
            parentNode: this._xfNode,
            tagName: "alignment",
            attributeName: "horizontal",
            args: arguments
        });
    }

    verticalAlignment(alignment) {
        debug("verticalAlignment(%o)", arguments);
        if (this._styles) return this._iterateStyles("verticalAlignment", arguments);
        return this._attributeValue({
            methodName: "verticalAlignment",
            parentNode: this._xfNode,
            tagName: "alignment",
            attributeName: "vertical",
            args: arguments
        });
    }

    wrappedText(wrappedText) {
        debug("wrappedText(%o)", arguments);
        if (this._styles) return this._iterateStyles("wrappedText", arguments);
        return this._attributeValue({
            methodName: "wrappedText",
            parentNode: this._xfNode,
            tagName: "alignment",
            attributeName: "wrapText",
            args: arguments,
            toStringConverter: val => val ? "1" : "0",
            fromStringConverter: val => val === "1"
        });
    }

    indent(indent) {
        debug("indent(%o)", arguments);
        if (this._styles) return this._iterateStyles("indent", arguments);
        return this._attributeValue({
            methodName: "indent",
            parentNode: this._xfNode,
            tagName: "alignment",
            attributeName: "indent",
            args: arguments,
            fromStringConverter: val => parseInt(val)
        });
    }

    textRotation(indent) {
        debug("textRotation(%o)", arguments);
        if (this._styles) return this._iterateStyles("textRotation", arguments);
        return this._attributeValue({
            methodName: "textRotation",
            parentNode: this._xfNode,
            tagName: "alignment",
            attributeName: "textRotation",
            args: arguments,
            fromStringConverter: val => parseInt(val)
        });
    }

    angleTextCounterclockwise() {
        debug("angleTextCounterclockwise(%o)", arguments);
        if (this._styles) return this._iterateStyles("angleTextCounterclockwise", arguments);
        return this._shortcutValue({
            methodName: "angleTextCounterclockwise",
            upstreamMethodName: "textRotation",
            value: 45,
            args: arguments
        });
    }

    angleTextClockwise() {
        debug("angleTextClockwise(%o)", arguments);
        if (this._styles) return this._iterateStyles("angleTextClockwise", arguments);
        return this._shortcutValue({
            methodName: "angleTextClockwise",
            upstreamMethodName: "textRotation",
            value: 135,
            args: arguments
        });
    }

    verticalText() {
        debug("verticalText(%o)", arguments);
        if (this._styles) return this._iterateStyles("verticalText", arguments);
        return this._shortcutValue({
            methodName: "verticalText",
            upstreamMethodName: "textRotation",
            value: 255,
            args: arguments
        });
    }

    rotateTextUp() {
        debug("rotateTextUp(%o)", arguments);
        if (this._styles) return this._iterateStyles("rotateTextUp", arguments);
        return this._shortcutValue({
            methodName: "rotateTextUp",
            upstreamMethodName: "textRotation",
            value: 90,
            args: arguments
        });
    }

    rotateTextDown() {
        debug("rotateTextDown(%o)", arguments);
        if (this._styles) return this._iterateStyles("rotateTextDown", arguments);
        return this._shortcutValue({
            methodName: "rotateTextDown",
            upstreamMethodName: "textRotation",
            value: 180,
            args: arguments
        });
    }

    _borderStyle(side, args) {
        debug("_borderStyle(%s, ...)", side);
        if (this._styles) return this._iterateStyles(`${side}BorderStyle`, arguments);
        return this._attributeValue({
            methodName: `${side}BorderStyle`,
            parentNode: this._borderNode,
            tagName: side,
            attributeName: "style",
            args
        });
    }

    topBorderStyle() {
        debug("topBorderStyle(%o)", arguments);
        return this._borderStyle("top", arguments);
    }

    bottomBorderStyle() {
        debug("bottomBorderStyle(%o)", arguments);
        return this._borderStyle("bottom", arguments);
    }

    leftBorderStyle() {
        debug("leftBorderStyle(%o)", arguments);
        return this._borderStyle("left", arguments);
    }

    rightBorderStyle() {
        debug("rightBorderStyle(%o)", arguments);
        return this._borderStyle("right", arguments);
    }
}

module.exports = Style;
