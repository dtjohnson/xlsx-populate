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

    _get_italic() {
        debug("_get_italic(%o)", arguments);
        return jq.has(this._fontNode, "i[0]");
    }

    _set_italic(italic) {
        debug("_set_italic(%o)", arguments);
        if (italic) jq.setIfNeeded(this._fontNode, "i[0]", {});
        else delete this._fontNode.i;
    }

    _get_underline() {
        debug("_get_underline(%o)", arguments);
        const val = jq.get(this._fontNode, "u[0].$.val");
        if (val) return val;
        return jq.has(this._fontNode, "u[0]");
    }

    _set_underline(underline) {
        debug("_set_underline(%o)", arguments);
        if (typeof underline === "string") jq.set(this._fontNode, "u[0].$.val", underline);
        else if (underline) jq.setIfNeeded(this._fontNode, "u[0]", {});
        else delete this._fontNode.u;
    }

    _get_strikethrough() {
        debug("_get_strikethrough(%o)", arguments);
        return jq.has(this._fontNode, "strike[0]");
    }

    _set_strikethrough(strikethrough) {
        debug("_set_strikethrough(%o)", arguments);
        if (strikethrough) jq.setIfNeeded(this._fontNode, "strike[0]", {});
        else delete this._fontNode.strike;
    }

    _getFontVerticalAlignment() {
        debug("_getFontVerticalAlignment(%o)", arguments);
        return jq.get(this._fontNode, "vertAlign[0].$.val");
    }

    _setFontVerticalAlignment(alignment) {
        debug("_setFontVerticalAlignment(%o)", arguments);
        jq.set(this._fontNode, "vertAlign[0].$.val", alignment);
        if (jq.isEmpty(this._fontNode, "vertAlign[0].$")) delete this._fontNode.vertAlign;
    }

    _get_superscript() {
        debug("_get_superscript(%o)", arguments);
        return this._getFontVerticalAlignment() === "superscript";
    }

    _set_superscript(superscript) {
        debug("_set_superscript(%o)", arguments);
        this._setFontVerticalAlignment(superscript ? "superscript" : null);
    }

    _get_subscript() {
        debug("_get_subscript(%o)", arguments);
        return this._getFontVerticalAlignment() === "subscript";
    }

    _set_subscript(subscript) {
        debug("_set_subscript(%o)", arguments);
        this._setFontVerticalAlignment(subscript ? "subscript" : null);
    }

    _get_fontSize() {
        debug("_get_fontSize(%o)", arguments);
        return jq.get(this._fontNode, "sz[0].$.val");
    }

    _set_fontSize(size) {
        debug("_set_fontSize(%o)", arguments);
        jq.set(this._fontNode, "sz[0].$.val", size);
        if (jq.isEmpty(this._fontNode, "sz[0].$")) delete this._fontNode.sz;
    }

    _get_fontFamily() {
        debug("_get_fontFamily(%o)", arguments);
        return jq.get(this._fontNode, "name[0].$.val");
    }

    _set_fontFamily(family) {
        debug("_set_fontFamily(%o)", arguments);
        jq.set(this._fontNode, "name[0].$.val", family);
        if (jq.isEmpty(this._fontNode, "name[0].$")) delete this._fontNode.name;
    }

    _get_fontColor() {
        debug("_get_fontColor(%o)", arguments);
        return jq.apply(this._fontNode, "color[0].$", $ => $ && ($.rgb || $.indexed));
    }

    _set_fontColor(color) {
        debug("_set_fontColor(%o)", arguments);
        let rgb, indexed;
        if (typeof color === "string") rgb = color;
        else if (color >= 0) indexed = color;

        jq.set(this._fontNode, {
            "color[0].$.rgb": rgb,
            "color[0].$.indexed": indexed
        });

        if (jq.isEmpty(this._fontNode, "color[0].$")) delete this._fontNode.color;
    }

    _get_horizontalAlignment() {
        debug("_get_horizontalAlignment(%o)", arguments);
        return jq.get(this._xfNode, "alignment[0].$.horizontal");
    }

    _set_horizontalAlignment(alignment) {
        debug("_set_horizontalAlignment(%o)", arguments);
        jq.set(this._xfNode, "alignment[0].$.horizontal", alignment);
        if (jq.isEmpty(this._xfNode, "alignment[0].$")) delete this._xfNode.alignment;
    }

    _get_verticalAlignment() {
        debug("_get_verticalAlignment(%o)", arguments);
        return jq.get(this._xfNode, "alignment[0].$.vertical");
    }

    _set_verticalAlignment(alignment) {
        debug("_set_verticalAlignment(%o)", arguments);
        jq.set(this._xfNode, "alignment[0].$.vertical", alignment);
        if (jq.isEmpty(this._xfNode, "alignment[0].$")) delete this._xfNode.alignment;
    }

    _get_wrappedText() {
        debug("_get_wrappedText(%o)", arguments);
        return jq.get(this._xfNode, "alignment[0].$.wrapText");
    }

    _set_wrappedText(wrappedText) {
        debug("_set_wrappedText(%o)", arguments);
        jq.set(this._xfNode, "alignment[0].$.wrapText", wrappedText);
        if (jq.isEmpty(this._xfNode, "alignment[0].$")) delete this._xfNode.alignment;
    }

    _get_indent() {
        debug("_get_indent(%o)", arguments);
        return jq.get(this._xfNode, "alignment[0].$.indent");
    }

    _set_indent(indent) {
        debug("_set_indent(%o)", arguments);
        jq.set(this._xfNode, "alignment[0].$.indent", indent);
        if (jq.isEmpty(this._xfNode, "alignment[0].$")) delete this._xfNode.alignment;
    }

    // TODO: Negative values?
    _get_textRotation() {
        debug("_get_textRotation(%o)", arguments);
        return jq.get(this._xfNode, "alignment[0].$.textRotation");
    }

    _set_textRotation(textRotation) {
        debug("_set_textRotation(%o)", arguments);
        jq.set(this._xfNode, "alignment[0].$.textRotation", textRotation);
        if (jq.isEmpty(this._xfNode, "alignment[0].$")) delete this._xfNode.alignment;
    }

    _get_angleTextCounterclockwise() {
        debug("_get_angleTextCounterclockwise(%o)", arguments);
        return this._get_textRotation() === 45;
    }

    _set_angleTextCounterclockwise(value) {
        debug("_set_angleTextCounterclockwise(%o)", arguments);
        this._set_textRotation(value ? 45 : null);
    }

    _get_angleTextClockwise() {
        debug("_get_angleTextClockwise(%o)", arguments);
        return this._get_textRotation() === 135;
    }

    _set_angleTextClockwise(value) {
        debug("_set_angleTextClockwise(%o)", arguments);
        this._set_textRotation(value ? 135 : null);
    }

    _get_verticalText() {
        debug("_get_verticalText(%o)", arguments);
        return this._get_textRotation() === 255;
    }

    _set_verticalText(value) {
        debug("_set_verticalText(%o)", arguments);
        this._set_textRotation(value ? 255 : null);
    }

    _get_rotateTextUp() {
        debug("__rotateTextUp(%o)", arguments);
        return this._get_textRotation() === 90;
    }

    _set_rotateTextUp(value) {
        debug("_set_rotateTextUp(%o)", arguments);
        this._set_textRotation(value ? 90 : null);
    }

    _get_rotateTextDown() {
        debug("_get_rotateTextDown(%o)", arguments);
        return this._get_textRotation() === 180;
    }

    _set_rotateTextDown(value) {
        debug("_set_rotateTextDown(%o)", arguments);
        this._set_textRotation(value ? 180 : null);
    }

    _get_numberFormat() {
        debug("_get_numberFormat(%o)", arguments);
        const numFmtId = jq.get(this._xfNode, '$.numFmtId') || 0;
        return this._styleSheet.getNumberFormatCode(numFmtId);
    }

    _set_numberFormat(formatCode) {
        debug("_set_numberFormat(%o)", arguments);
        jq.set(this._xfNode, '$.numFmtId', this._styleSheet.getNumberFormatId(formatCode));
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

    _get_diagonalBorderDirection() {
        return jq.get(this._getBorder(), "diagonal.direction");
    }

    _set_diagonalBorderDirection(direction) {
        this._setBorder({ diagonal: { direction } });
    }
}

["left", "right", "top", "bottom", "diagonal"].forEach(side => {
    _Style.prototype[`_get_${side}Border`] = function () {
        return jq.get(this._getBorder(), side);
    };

    _Style.prototype[`_set_${side}Border`] = function (settings) {
        this._setBorder({ [side]: settings });
    };

    _Style.prototype[`_get_${side}BorderColor`] = function () {
        return jq.get(this._getBorder(), `${side}.color`);
    };

    _Style.prototype[`_set_${side}BorderColor`] = function (color) {
        this._setBorder({ [side]: { color } });
    };

    _Style.prototype[`_get_${side}BorderStyle`] = function () {
        return jq.get(this._getBorder(), `${side}.style`);
    };

    _Style.prototype[`_set_${side}BorderStyle`] = function (style) {
        this._setBorder({ [side]: { style } });
    };
});

module.exports = _Style;
