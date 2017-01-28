"use strict";

const _ArgParser = require("./_ArgParser");
const debug = require("./debug")("_Style");
const xq = require("./xq");

// TODO: Double underline.

class Style {
    constructor(styleSheet, id, xfNode, fontNode, borderNode) {
        this._styleSheet = styleSheet;
        this._id = id;
        this._xfNode = xfNode;
        this._fontNode = fontNode;
        this._borderNode = borderNode;
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
        if (arguments.length === 0) {
            const result = xq.query(this._fontNode, {
                u: {
                    "@val": { $type: String, $optional: true }
                }
            });

            if (!result) return false;
            if (result.u['@val'] === "double") return "double";
            return true;
        }

        const update = { u: null };
        if (underline) update.u = {};
        if (underline === "double") update.u["@val"] = "double";
        xq.update(this._fontNode, update);
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
            }, true);
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
            }, true);
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
            }, true);
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
            }, true);
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
            }, true);
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
            }, true);
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
            }, true);
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
            }, true);
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

    __numberFormat(formatCode) {
        debug("__numberFormat(%o)", arguments);
        if (arguments.length === 0) {
            const numFmtId = xq.query(this._xfNode, {
                '@numFmtId': Number
            }, true) || 0;

            return this._styleSheet.getNumberFormatCode(numFmtId);
        }

        xq.update(this._xfNode, {
            '@numFmtId': this._styleSheet.getNumberFormatId(formatCode)
        });
        return this;
    }

    // TODO: Consider various border options. Should we merge like CSS borders?

    _borderStyle(side, args) {
        debug("_borderStyle(%o)", arguments);
        if (args.length === 0) {
            return xq.query(this._borderNode, {
                [side]: {
                    "@style": String
                }
            }, true);
        }

        xq.update(this._borderNode, {
            [side]: {
                "@style": args[0] || null,
                $removeIfEmpty: true
            }
        });

        return this;
    }

    __topBorderStyle() {
        debug("__topBorderStyle(%o)", arguments);
        return this._borderStyle("top", arguments);
    }

    __bottomBorderStyle() {
        debug("__bottomBorderStyle(%o)", arguments);
        return this._borderStyle("bottom", arguments);
    }

    __leftBorderStyle() {
        debug("leftBorderStyle(%o)", arguments);
        return this._borderStyle("left", arguments);
    }

    __rightBorderStyle() {
        debug("rightBorderStyle(%o)", arguments);
        return this._borderStyle("right", arguments);
    }
}

module.exports = Style;

/*

 Fill:
 style with applyFill="1" -> fill with:
 <patternFill patternType="solid">
 <fgColor rgb="FFFFFF00"/>
 <bgColor indexed="64"/>
 </patternFill>

 Font color x:
 style -> font with <color rgb="x"/> indexed, theme, or rgb

 Border:
 style with applyBorder="1" -> border:
 <border>
 <left /> style="thin|medium|thick|double"
 child <color> indexed = 64 is black (not sure if necessary)
 ...
 </border>

 Number Format:
 style -> numFmtId
 Create numFmt if not standard: http://polymathprogrammer.com/2011/02/15/built-in-styles-for-excel-open-xml/
 */
