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
                "@val": alignment || null,
                $removeIfEmpty: true
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
                "@val": size || null,
                $removeIfEmpty: true
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
                "@val": family || null,
                $removeIfEmpty: true
            }
        });
    }

    // TODO: # prefix?, rgb(x, x, x)?
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
                "@rgb": rgb || null,
                "@indexed": indexed || null,
                $removeIfEmpty: true
            }
        });
    }

    __horizontalAlignment(alignment) {
        debug("__horizontalAlignment(%o)", arguments);
        if (arguments.length === 0) {
            return xq.query(this._xfNode, {
                alignment: {
                    "@horizontal": String
                }
            }, "alignment.@horizontal");
        }

        xq.update(this._xfNode, {
            alignment: {
                "@horizontal": alignment || null,
                $removeIfEmpty: true
            }
        });
    }

    __verticalAlignment(alignment) {
        debug("__verticalAlignment(%o)", arguments);
        if (arguments.length === 0) {
            return xq.query(this._xfNode, {
                alignment: {
                    "@vertical": String
                }
            }, "alignment.@vertical");
        }

        xq.update(this._xfNode, {
            alignment: {
                "@vertical": alignment || null,
                $removeIfEmpty: true
            }
        });
    }

    __wrappedText(wrappedText) {
        debug("__wrappedText(%o)", arguments);
        if (arguments.length === 0) {
            return !!xq.query(this._xfNode, {
                alignment: {
                    "@wrapText": Boolean
                }
            }, "alignment.@wrapText");
        }

        xq.update(this._xfNode, {
            alignment: {
                "@wrapText": wrappedText || null,
                $removeIfEmpty: true
            }
        });
    }

    __indent(indent) {
        debug("__indent(%o)", arguments);
        if (arguments.length === 0) {
            return xq.query(this._xfNode, {
                alignment: {
                    "@indent": Number
                }
            }, "alignment.@indent");
        }

        xq.update(this._xfNode, {
            alignment: {
                "@indent": indent || null,
                $removeIfEmpty: true
            }
        });
    }

    // TODO: Negative values?
    __textRotation(textRotation) {
        debug("__textRotation(%o)", arguments);
        if (arguments.length === 0) {
            return xq.query(this._xfNode, {
                alignment: {
                    "@textRotation": Number
                }
            }, "alignment.@textRotation");
        }

        xq.update(this._xfNode, {
            alignment: {
                "@textRotation": textRotation || null,
                $removeIfEmpty: true
            }
        });
    }

    __angleTextCounterclockwise(value) {
        debug("__angleTextCounterclockwise(%o)", arguments);
        if (arguments.length === 0) return this.__textRotation() === 45;
        return this.__textRotation(value ? 45 : null);
    }

    __angleTextClockwise(value) {
        debug("__angleTextClockwise(%o)", arguments);
        if (arguments.length === 0) return this.__textRotation() === 135;
        return this.__textRotation(value ? 135 : null);
    }

    __verticalText(value) {
        debug("__verticalText(%o)", arguments);
        if (arguments.length === 0) return this.__textRotation() === 255;
        return this.__textRotation(value ? 255 : null);
    }

    __rotateTextUp(value) {
        debug("__rotateTextUp(%o)", arguments);
        if (arguments.length === 0) return this.__textRotation() === 90;
        return this.__textRotation(value ? 90 : null);
    }

    __rotateTextDown(value) {
        debug("__rotateTextDown(%o)", arguments);
        if (arguments.length === 0) return this.__textRotation() === 180;
        return this.__textRotation(value ? 180 : null);
    }


    // TODO: Fix these below
    // TODO: Consider various border options. Should we merge like CSS borders?

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
