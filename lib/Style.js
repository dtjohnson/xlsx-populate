"use strict";

const _ArgParser = require("./_ArgParser");
const debug = require("./debug")("_Style");
const xq = require("./xq");

class Style {
    constructor() {
        this._id = arguments[0];
        this._xfNode = arguments[1];
        this._fontNode = arguments[2];
        this._borderNode = arguments[3];
    }

    style() {
        debug("style(%o)", arguments);
        return new _ArgParser("_Style.style")
            .case(String, name => {
                const methodName = `__${name}`;
                if (!this[methodName]) throw new Error(`_Style.style: '${name}' is not a valid style`);
                return this[methodName]();
            })
            .case([String, undefined], (name, value) => {
                const methodName = `__${name}`;
                if (!this[methodName]) throw new Error(`_Style.style: '${name}' is not a valid style`);
                this[methodName](value);
                return this;
            })
            .parse(arguments);
    }

    __bold(bold) {
        debug("__bold(%o)", arguments);
        if (arguments.length === 0) return !!xq.query(this._fontNode, { b: {} });
        xq.update(this._fontNode, { b: bold ? {} : null });
    }

    __italic(italic) {
        debug("__italic(%o)", arguments);
        if (arguments.length === 0) return !!xq.query(this._fontNode, { i: {} });
        xq.update(this._fontNode, { i: italic ? {} : null });
    }

    __underline(underline) {
        debug("__underline(%o)", arguments);
        if (arguments.length === 0) return !!xq.query(this._fontNode, { u: {} });
        xq.update(this._fontNode, { u: underline ? {} : null });
    }

    __strikethrough(strikethrough) {
        debug("__strikethrough(%o)", arguments);
        if (arguments.length === 0) return !!xq.query(this._fontNode, { strike: {} });
        xq.update(this._fontNode, { strike: strikethrough ? {} : null });
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

    __fontVerticalAlignment(alignment) {
        debug("__fontVerticalAlignment(%o)", arguments);
        if (arguments.length === 0) {
            return xq.query(this._fontNode, {
                vertAlign: {
                    "@val": String
                }
            }, "vertAlign.@val");
        }

        xq.update(this._fontNode, {
            vertAlign: {
                "@val": alignment
            }
        });
    }

    __superscript(superscript) {
        debug("__superscript(%o)", arguments);
        if (arguments.length === 0) return this.__fontVerticalAlignment() === "superscript";
        return this.__fontVerticalAlignment(superscript ? "superscript" : null);
    }

    __subscript(subscript) {
        debug("__subscript(%o)", arguments);
        if (arguments.length === 0) return this.__fontVerticalAlignment() === "subscript";
        return this.__fontVerticalAlignment(subscript ? "subscript" : null);
    }

    __fontSize(size) {
        debug("__fontSize(%o)", arguments);
        if (arguments.length === 0) {
            return xq.query(this._fontNode, {
                sz: {
                    "@val": Number
                }
            }, "sz.@val");
        }

        xq.update(this._fontNode, {
            sz: {
                "@val": size
            }
        });
    }

    __fontFamily(family) {
        debug("__fontFamily(%o)", arguments);
        if (arguments.length === 0) {
            return xq.query(this._fontNode, {
                name: {
                    "@val": String
                }
            }, "name.@val");
        }

        xq.update(this._fontNode, {
            name: {
                "@val": family
            }
        });
    }

    __fontColor(color) {
        debug("__fontColor(%o)", arguments);
        if (arguments.length === 0) {
            const result = xq.query(this._fontNode, {
                color: {
                    "@rgb": { $type: String, $optional: true },
                    "@indexed": { $type: Number, $optional: true }
                }
            });

            return result && (result.color['@rgb'] || result.color['@indexed']);
        }

        let rgb = null, indexed = null;
        if (typeof color === "string") rgb = color;
        else indexed = color;

        xq.update(this._fontNode, {
            color: {
                "@rgb": rgb,
                "@indexed": indexed
            }
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
