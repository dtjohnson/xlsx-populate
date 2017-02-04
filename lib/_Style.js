"use strict";

// TODO: Cell backgrounds: solid, patterned, gradients
// TODO: Font: shrink to fit
// TODO: Horizontal alignment: General, Fill, Justify, Center Across Selection, Distributed
// TODO: Vertical alignment: Distributed, Justify Distributed
// TODO: Text direction: Content, Left-to-Right, Right-to-Left

/* eslint camelcase:off */

const _ArgParser = require("./_ArgParser");
const debug = require("./debug")("_Style");
const jq = require("./jq");
const _ = require("lodash");

class _Style {
    constructor(styleSheet, id, xfNode, fontNode, fillNode, borderNode) {
        this._styleSheet = styleSheet;
        this._id = id;
        this._xfNode = xfNode;
        this._fontNode = fontNode;
        this._fillNode = fillNode;
        this._borderNode = borderNode;
    }

    style() {
        debug("style(%o)", arguments);
        return new _ArgParser("_Style.style")
            .case(String, name => {
                const getterName = `_get_${name}`;
                if (!this[getterName]) throw new Error(`_Style.style: '${name}' is not a valid style`);
                return this[getterName]();
            })
            .case([String, undefined], (name, value) => {
                const setterName = `_set_${name}`;
                if (!this[setterName]) throw new Error(`_Style.style: '${name}' is not a valid style`);
                this[setterName](value);
                return this;
            })
            .parse(arguments);
    }

    _get_bold() {
        debug("_get_bold(%o)", arguments);
        return _.has(this._fontNode, "b[0]");
    }

    _set_bold(bold) {
        debug("_set_bold(%o)", arguments);
        if (bold) jq.setIfNeeded(this._fontNode, "b[0]", {});
        else delete this._fontNode.b;
    }

    __italic(italic) {
        debug("__italic(%o)", arguments);
        if (arguments.length === 0) return jq.has(this._fontNode, "i[0]");
        if (italic) jq.setIfNeeded(this._fontNode, "i[0]", {});
        else delete this._fontNode.i;
    }

    __underline(underline) {
        debug("__underline(%o)", arguments);
        if (arguments.length === 0) {
            const val = jq.get(this._fontNode, "u[0].$.val");
            if (val) return val;
            return jq.has(this._fontNode, "u[0]");
        }

        if (typeof underline === "string") jq.set(this._fontNode, "u[0].$.val", underline);
        else if (underline) jq.setIfNeeded(this._fontNode, "strike[0]", {});
        else delete this._fontNode.u;
    }

    __strikethrough(strikethrough) {
        debug("__strikethrough(%o)", arguments);
        if (arguments.length === 0) return jq.has(this._fontNode, "strike[0]");
        if (strikethrough) jq.setIfNeeded(this._fontNode, "strike[0]", {});
        else delete this._fontNode.strike;
    }

    __fontVerticalAlignment(alignment) {
        debug("__fontVerticalAlignment(%o)", arguments);
        if (arguments.length === 0) return jq.get(this._fontNode, "vertAlign[0].$.val");
        jq.set(this._fontNode, "vertAlign[0].$.val", alignment);
        if (jq.isEmpty(this._fontNode, "vertAlign[0].$")) delete this._fontNode.vertAlign;
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
        if (arguments.length === 0) return jq.get(this._fontNode, "sz[0].$.val");
        jq.set(this._fontNode, "sz[0].$.val", size);
        if (jq.isEmpty(this._fontNode, "sz[0].$")) delete this._fontNode.sz;
    }

    __fontFamily(family) {
        debug("__fontFamily(%o)", arguments);
        if (arguments.length === 0) return jq.get(this._fontNode, "name[0].$.val");
        jq.set(this._fontNode, "name[0].$.val", family);
        if (jq.isEmpty(this._fontNode, "name[0].$")) delete this._fontNode.name;
    }

    // TODO: # prefix?, rgb(x, x, x)?
    __fontColor(color) {
        debug("__fontColor(%o)", arguments);
        if (arguments.length === 0) {
            return jq.apply(this._fontNode, "color[0].$", $ => $ && ($.rgb || $.indexed));
        }

        let rgb, indexed;
        if (typeof color === "string") rgb = color;
        else if (color >= 0) indexed = color;

        jq.set(this._fontNode, {
            "color[0].$.rgb": rgb,
            "color[0].$.indexed": indexed
        });

        if (jq.isEmpty(this._fontNode, "color[0].$")) delete this._fontNode.color;
    }

    __horizontalAlignment(alignment) {
        debug("__horizontalAlignment(%o)", arguments);
        if (arguments.length === 0) return jq.get(this._xfNode, "alignment[0].$.horizontal");
        jq.set(this._xfNode, "alignment[0].$.horizontal", alignment);
        if (jq.isEmpty(this._xfNode, "alignment[0].$")) delete this._xfNode.alignment;
    }

    __verticalAlignment(alignment) {
        debug("__verticalAlignment(%o)", arguments);
        if (arguments.length === 0) return jq.get(this._xfNode, "alignment[0].$.vertical");
        jq.set(this._xfNode, "alignment[0].$.vertical", alignment);
        if (jq.isEmpty(this._xfNode, "alignment[0].$")) delete this._xfNode.alignment;
    }

    __wrappedText(wrappedText) {
        debug("__wrappedText(%o)", arguments);
        if (arguments.length === 0) return jq.get(this._xfNode, "alignment[0].$.wrapText");
        jq.set(this._xfNode, "alignment[0].$.wrapText", wrappedText);
        if (jq.isEmpty(this._xfNode, "alignment[0].$")) delete this._xfNode.alignment;
    }

    __indent(indent) {
        debug("__indent(%o)", arguments);
        if (arguments.length === 0) return jq.get(this._xfNode, "alignment[0].$.indent");
        jq.set(this._xfNode, "alignment[0].$.indent", indent);
        if (jq.isEmpty(this._xfNode, "alignment[0].$")) delete this._xfNode.alignment;
    }

    // TODO: Negative values?
    __textRotation(textRotation) {
        debug("__textRotation(%o)", arguments);
        if (arguments.length === 0) return jq.get(this._xfNode, "alignment[0].$.textRotation");
        jq.set(this._xfNode, "alignment[0].$.textRotation", textRotation);
        if (jq.isEmpty(this._xfNode, "alignment[0].$")) delete this._xfNode.alignment;
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
            const numFmtId = jq.get(this._xfNode, '$.numFmtId') || 0;
            return this._styleSheet.getNumberFormatCode(numFmtId);
        }

        jq.set(this._xfNode, '$.numFmtId', this._styleSheet.getNumberFormatId(formatCode));
    }

    // TODO: Consider various border options. Should we merge like CSS borders?

    _applyBorder() {

    }

    _getBorder() {
        const result = {};
        ["left", "right", "top", "bottom", "diagonal"].forEach(side => {
            jq.set(result, `${side}.style`, jq.get(this._borderNode, `${side}[0].$.style`));
            jq.set(result, `${side}.color`, jq.apply(this._borderNode, `${side}[0].color[0].$`, $ => $ && ($.rgb || $.indexed)));
            if (side === "diagonal") {
                const up = jq.get(this._borderNode, "$.diagonalUp");
                const down = jq.get(this._borderNode, "$.diagonalUp");
                let direction;
                if (up && down) direction = "both";
                else if (up) direction = "up";
                else if (down) direction = "down";
                jq.set(result, 'diagonal.direction', direction);
            }
        });

        return result;
    }

    _setBorder(settings) {
        _.forOwn(settings, (setting, side) => {
            if (typeof setting === "boolean") {
                setting = { style: setting ? "thin" : null };
            } else if (typeof setting === "string") {
                setting = { style: setting };
            } else if (setting === null || setting === undefined) {
                setting = { style: null, color: null, direction: null };
            }

            if (setting.hasOwnProperty("style")) {
                jq.set(this._borderNode, `${side}[0].$.style`, setting.style);
            }

            if (setting.hasOwnProperty("color")) {
                let rgb, indexed;
                if (typeof setting.color === "string") rgb = setting.color;
                else if (setting.color >= 0) indexed = setting.color;
                jq.set(this._borderNode, {
                    [`${side}[0].color[0].$.rgb`]: rgb,
                    [`${side}[0].color[0].$.indexed`]: indexed
                });
            }

            if (side === "diagonal") {
                jq.set(this._borderNode, "$.diagonalUp", setting.direction === "up" || setting.direction === "both" ? 1 : null);
                jq.set(this._borderNode, "$.diagonalDown", setting.direction === "down" || setting.direction === "both" ? 1 : null);
            }
        });
    }

    _get_border() {
        const borders = this._getBorder();
        return _.isEmpty(borders) ? null : borders;
    }

    _set_border(settings) {
        if (_.isObject(settings)) {
            settings = _.defaults(settings, {
                left: null,
                right: null,
                top: null,
                bottom: null,
                diagonal: null
            });
            this._setBorder(settings);
        } else {
            this._setBorder({
                left: settings,
                right: settings,
                top: settings,
                bottom: settings
            });
        }
    }

    _get_borderColor() {
        return _.mapValues(this._getBorder(), value => value.color);
    }

    _set_borderColor(color) {
        if (_.isObject(color)) {
            this._setBorder(_.mapValues(color, color => ({ color })));
        } else {
            this._setBorder({
                left: { color },
                right: { color },
                top: { color },
                bottom: { color },
                diagonal: { color }
            });
        }
    }

    _get_borderStyle() {
        return _.mapValues(this._getBorder(), value => value.style);
    }

    _set_borderStyle(style) {
        if (_.isObject(style)) {
            this._setBorder(_.mapValues(style, style => ({ style })));
        } else {
            this._setBorder({
                left: { style },
                right: { style },
                top: { style },
                bottom: { style }
            });
        }
    }

    _get_borderDiagonalDirection() {
        return jq.get(this._getBorder(), "diagonal.direction");
    }

    _set_borderDiagonalDirection(direction) {
        this._setBorder({ diagonal: { direction } });
    }
}

["left", "right", "top", "bottom", "diagonal"].forEach(side => {
    const sideUC = _.upperFirst(side);
    _Style.prototype[`_get_border${sideUC}`] = function () {
        return jq.get(this._getBorder(), side);
    };

    _Style.prototype[`_set_border${sideUC}`] = function (settings) {
        this._setBorder({ [side]: settings });
    };

    _Style.prototype[`_get_border${sideUC}Color`] = function () {
        return jq.get(this._getBorder(), `${side}.color`);
    };

    _Style.prototype[`_set_border${sideUC}Color`] = function (color) {
        this._setBorder({ [side]: { color } });
    };

    _Style.prototype[`_get_border${sideUC}Style`] = function () {
        return jq.get(this._getBorder(), `${side}.style`);
    };

    _Style.prototype[`_set_border${sideUC}Style`] = function (style) {
        this._setBorder({ [side]: { style } });
    };
});

module.exports = _Style;
