"use strict";

/* eslint camelcase:off */

const _ArgParser = require("./_ArgParser");
const debug = require("./debug")("_Style");
const jq = require("./jq");
const _ = require("lodash");
const xmlq = require("./xmlq");
const colorIndexes = require("./colorIndexes");

/**
 * A style.
 */
class _Style {
    /**
     * Creates a new instance of _Style.
     * @constructor
     * @param {_StyleSheet} styleSheet - The styleSheet.
     * @param {number} id - The style ID.
     * @param {{}} xfNode - The xf node.
     * @param {{}} fontNode - The font node.
     * @param {{}} fillNode - The fill node.
     * @param {{}} borderNode - The border node.
     */
    constructor(styleSheet, id, xfNode, fontNode, fillNode, borderNode) {
        this._styleSheet = styleSheet;
        this._id = id;
        this._xfNode = xfNode;
        this._fontNode = fontNode;
        this._fillNode = fillNode;
        this._borderNode = borderNode;
    }

    /**
     * Gets or sets a style.
     * @param {string} name - The style name.
     * @param {*} [value] - The value to set.
     * @returns {*|_Style} The value if getting or the style if setting.
     */
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

    _getColor(node, name) {
        const child = xmlq.findChild(node, name);
        if (!child || !child.attributes) return;

        const color = {};
        if (child.attributes.hasOwnProperty('rgb')) color.rgb = child.attributes.rgb;
        else if (child.attributes.hasOwnProperty('theme')) color.theme = child.attributes.theme;
        else if (child.attributes.hasOwnProperty('indexed')) color.rgb = colorIndexes[child.attributes.indexed];

        if (child.attributes.hasOwnProperty('tint')) color.tint = child.attributes.tint;

        if (_.isEmpty(color)) return;

        return color;
    }

    _setColor(node, name, color) {
        if (typeof color === "string") color = { rgb: color };
        else if (typeof color === "number") color = { theme: color };

        xmlq.setChildAttributes(node, name, {
            rgb: color && color.rgb && color.rgb.toUpperCase(),
            indexed: null,
            theme: color && color.theme,
            tint: color && color.tint
        });

        xmlq.removeChildIfEmpty(node, 'color');
    }

    _get_bold() {
        debug("_get_bold(%o)", arguments);
        return xmlq.hasChild(this._fontNode, 'b');
    }

    _set_bold(bold) {
        debug("_set_bold(%o)", arguments);
        if (bold) xmlq.appendChildIfNotFound(this._fontNode, "b");
        else xmlq.removeChild(this._fontNode, 'b');
    }

    _get_italic() {
        debug("_get_italic(%o)", arguments);
        return xmlq.hasChild(this._fontNode, 'i');
    }

    _set_italic(italic) {
        debug("_set_italic(%o)", arguments);
        if (italic) xmlq.appendChildIfNotFound(this._fontNode, "i");
        else xmlq.removeChild(this._fontNode, 'i');
    }

    _get_underline() {
        debug("_get_underline(%o)", arguments);
        const uNode = xmlq.findChild(this._fontNode, 'u');
        return uNode ? uNode.attributes.val || true : false;
    }

    _set_underline(underline) {
        debug("_set_underline(%o)", arguments);
        if (underline) {
            const uNode = xmlq.appendChildIfNotFound(this._fontNode, "u");
            const val = typeof underline === 'string' ? underline : null;
            xmlq.setAttributes(uNode, { val });
        } else {
            xmlq.removeChild(this._fontNode, 'u');
        }
    }

    _get_strikethrough() {
        debug("_get_strikethrough(%o)", arguments);
        return xmlq.hasChild(this._fontNode, 'strike');
    }

    _set_strikethrough(strikethrough) {
        debug("_set_strikethrough(%o)", arguments);
        if (strikethrough) xmlq.appendChildIfNotFound(this._fontNode, "strike");
        else xmlq.removeChild(this._fontNode, 'strike');
    }

    _getFontVerticalAlignment() {
        debug("_getFontVerticalAlignment(%o)", arguments);
        return xmlq.getChildAttribute(this._fontNode, 'vertAlign', "val");
    }

    _setFontVerticalAlignment(alignment) {
        debug("_setFontVerticalAlignment(%o)", arguments);
        xmlq.setChildAttributes(this._fontNode, 'vertAlign', { val: alignment });
        xmlq.removeChildIfEmpty(this._fontNode, 'vertAlign');
    }

    _get_subscript() {
        debug("_get_subscript(%o)", arguments);
        return this._getFontVerticalAlignment() === "subscript";
    }

    _set_subscript(subscript) {
        debug("_set_subscript(%o)", arguments);
        this._setFontVerticalAlignment(subscript ? "subscript" : null);
    }

    _get_superscript() {
        debug("_get_superscript(%o)", arguments);
        return this._getFontVerticalAlignment() === "superscript";
    }

    _set_superscript(superscript) {
        debug("_set_superscript(%o)", arguments);
        this._setFontVerticalAlignment(superscript ? "superscript" : null);
    }

    _get_fontSize() {
        debug("_get_fontSize(%o)", arguments);
        return xmlq.getChildAttribute(this._fontNode, 'sz', "val");
    }

    _set_fontSize(size) {
        debug("_set_fontSize(%o)", arguments);
        xmlq.setChildAttributes(this._fontNode, 'sz', { val: size });
        xmlq.removeChildIfEmpty(this._fontNode, 'sz');
    }

    _get_fontFamily() {
        debug("_get_fontFamily(%o)", arguments);
        return xmlq.getChildAttribute(this._fontNode, 'name', "val");
    }

    _set_fontFamily(family) {
        debug("_set_fontFamily(%o)", arguments);
        xmlq.setChildAttributes(this._fontNode, 'name', { val: family });
        xmlq.removeChildIfEmpty(this._fontNode, 'name');
    }

    _get_fontColor() {
        debug("_get_fontColor(%o)", arguments);
        return this._getColor(this._fontNode, "color");
    }

    _set_fontColor(color) {
        debug("_set_fontColor(%o)", arguments);
        this._setColor(this._fontNode, "color", color);
    }

    _get_horizontalAlignment() {
        debug("_get_horizontalAlignment(%o)", arguments);
        return xmlq.getChildAttribute(this._xfNode, 'alignment', "horizontal");
    }

    _set_horizontalAlignment(alignment) {
        debug("_set_horizontalAlignment(%o)", arguments);
        xmlq.setChildAttributes(this._xfNode, 'alignment', { horizontal: alignment });
        xmlq.removeChildIfEmpty(this._xfNode, 'alignment');
    }

    _get_justifyLastLine() {
        debug("_get_justifyLastLine(%o)", arguments);
        return xmlq.getChildAttribute(this._xfNode, 'alignment', "justifyLastLine") === 1;
    }

    _set_justifyLastLine(justifyLastLine) {
        debug("_set_justifyLastLine(%o)", arguments);
        xmlq.setChildAttributes(this._xfNode, 'alignment', { justifyLastLine: justifyLastLine ? 1 : null });
        xmlq.removeChildIfEmpty(this._xfNode, 'alignment');
    }

    _get_indent() {
        debug("_get_indent(%o)", arguments);
        return xmlq.getChildAttribute(this._xfNode, 'alignment', "indent");
    }

    _set_indent(indent) {
        debug("_set_indent(%o)", arguments);
        xmlq.setChildAttributes(this._xfNode, 'alignment', { indent });
        xmlq.removeChildIfEmpty(this._xfNode, 'alignment');
    }

    _get_verticalAlignment() {
        debug("_get_verticalAlignment(%o)", arguments);
        return xmlq.getChildAttribute(this._xfNode, 'alignment', "vertical");
    }

    _set_verticalAlignment(alignment) {
        debug("_set_verticalAlignment(%o)", arguments);
        xmlq.setChildAttributes(this._xfNode, 'alignment', { vertical: alignment });
        xmlq.removeChildIfEmpty(this._xfNode, 'alignment');
    }

    _get_wrapText() {
        debug("_get_wrapText(%o)", arguments);
        return xmlq.getChildAttribute(this._xfNode, 'alignment', "wrapText") === 1;
    }

    _set_wrapText(wrapText) {
        debug("_set_wrapText(%o)", arguments);
        xmlq.setChildAttributes(this._xfNode, 'alignment', { wrapText: wrapText ? 1 : null });
        xmlq.removeChildIfEmpty(this._xfNode, 'alignment');
    }

    _get_shrinkToFit() {
        debug("_get_shrinkToFit(%o)", arguments);
        return xmlq.getChildAttribute(this._xfNode, 'alignment', "shrinkToFit") === 1;
    }

    _set_shrinkToFit(shrinkToFit) {
        debug("_set_shrinkToFit(%o)", arguments);
        xmlq.setChildAttributes(this._xfNode, 'alignment', { shrinkToFit: shrinkToFit ? 1 : null });
        xmlq.removeChildIfEmpty(this._xfNode, 'alignment');
    }

    _get_textDirection() {
        debug("_get_textDirection(%o)", arguments);
        const readingOrder = xmlq.getChildAttribute(this._xfNode, 'alignment', "readingOrder");
        if (readingOrder === 1) return "left-to-right";
        if (readingOrder === 2) return "right-to-left";
        return readingOrder;
    }

    _set_textDirection(textDirection) {
        debug("_set_textDirection(%o)", arguments);
        let readingOrder;
        if (textDirection === "left-to-right") readingOrder = 1;
        else if (textDirection === "right-to-left") readingOrder = 2;
        xmlq.setChildAttributes(this._xfNode, 'alignment', { readingOrder });
        xmlq.removeChildIfEmpty(this._xfNode, 'alignment');
    }

    _getTextRotation() {
        debug("_getTextRotation(%o)", arguments);
        return xmlq.getChildAttribute(this._xfNode, 'alignment', "textRotation");
    }

    _setTextRotation(textRotation) {
        debug("_setTextRotation(%o)", arguments);
        xmlq.setChildAttributes(this._xfNode, 'alignment', { textRotation });
        xmlq.removeChildIfEmpty(this._xfNode, 'alignment');
    }

    _get_textRotation() {
        debug("_get_textRotation(%o)", arguments);
        let textRotation = this._getTextRotation();

        // Negative angles in Excel correspond to values > 90 in OOXML.
        if (textRotation > 90) textRotation = 90 - textRotation;
        return textRotation;
    }

    _set_textRotation(textRotation) {
        debug("_set_textRotation(%o)", arguments);

        // Negative angles in Excel correspond to values > 90 in OOXML.
        if (textRotation < 0) textRotation = 90 - textRotation;
        this._setTextRotation(textRotation);
    }

    _get_angleTextCounterclockwise() {
        debug("_get_angleTextCounterclockwise(%o)", arguments);
        return this._getTextRotation() === 45;
    }

    _set_angleTextCounterclockwise(value) {
        debug("_set_angleTextCounterclockwise(%o)", arguments);
        this._setTextRotation(value ? 45 : null);
    }

    _get_angleTextClockwise() {
        debug("_get_angleTextClockwise(%o)", arguments);
        return this._getTextRotation() === 135;
    }

    _set_angleTextClockwise(value) {
        debug("_set_angleTextClockwise(%o)", arguments);
        this._setTextRotation(value ? 135 : null);
    }

    _get_rotateTextUp() {
        debug("__rotateTextUp(%o)", arguments);
        return this._getTextRotation() === 90;
    }

    _set_rotateTextUp(value) {
        debug("_set_rotateTextUp(%o)", arguments);
        this._setTextRotation(value ? 90 : null);
    }

    _get_rotateTextDown() {
        debug("_get_rotateTextDown(%o)", arguments);
        return this._getTextRotation() === 180;
    }

    _set_rotateTextDown(value) {
        debug("_set_rotateTextDown(%o)", arguments);
        this._setTextRotation(value ? 180 : null);
    }

    _get_verticalText() {
        debug("_get_verticalText(%o)", arguments);
        return this._getTextRotation() === 255;
    }

    _set_verticalText(value) {
        debug("_set_verticalText(%o)", arguments);
        this._setTextRotation(value ? 255 : null);
    }

    _get_fill() {
        const patternFillNode = xmlq.findChild(this._fillNode, 'patternFill');// jq.get(this._fillNode, "patternFill[0]");
        const gradientFillNode = xmlq.findChild(this._fillNode, 'gradientFill');// jq.get(this._fillNode, "gradientFill[0]");
        const patternType = patternFillNode && patternFillNode.attributes.patternType;// jq.get(patternFillNode, "$.patternType");

        if (patternType === "solid") {
            return {
                type: "solid",
                color: this._getColor(patternFillNode, "fgColor")
            };
        }

        if (patternType) {
            return {
                type: "pattern",
                pattern: patternType,
                foreground: this._getColor(patternFillNode, "fgColor"),
                background: this._getColor(patternFillNode, "bgColor")
            };
        }

        if (gradientFillNode) {
            const gradientType = gradientFillNode.attributes.type || "linear";
            const fill = {
                type: "gradient",
                gradientType,
                stops: _.map(gradientFillNode.children, stop => ({
                    position: stop.attributes.position,
                    color: this._getColor(stop, "color")
                }))
            };

            if (gradientType === "linear") {
                fill.angle = gradientFillNode.attributes.degree;// jq.get(gradientFillNode, "$.degree");
            } else {
                fill.left = gradientFillNode.attributes.left;//jq.get(gradientFillNode, "$.left");
                fill.right = gradientFillNode.attributes.right;//jq.get(gradientFillNode, "$.right");
                fill.top = gradientFillNode.attributes.top;//jq.get(gradientFillNode, "$.top");
                fill.bottom = gradientFillNode.attributes.bottom;//jq.get(gradientFillNode, "$.bottom");
            }

            return fill;
        }
    }

    _set_fill(fill) {
        this._fillNode.children = [];
        // delete this._fillNode.patternFill;
        // delete this._fillNode.gradientFill;

        // No fill
        if (_.isNil(fill)) return;

        // Pattern fill
        if (fill.type === "pattern") {
            const patternFill = {
                name: 'patternFill',
                attributes: { patternType: fill.pattern },
                children: []
            };
            this._fillNode.children.push(patternFill);
            // this._fillNode.patternFill = [patternFill];
            this._setColor(patternFill, "fgColor", fill.foreground);
            this._setColor(patternFill, "bgColor", fill.background);
            // jq.set(patternFill, {
            //     "$.patternType": fill.pattern
            // });

            return;
        }

        // Gradient fill
        if (fill.type === "gradient") {
            const gradientFill = { name: 'gradientFill', attributes: {}, children: [] };
            this._fillNode.children.push(gradientFill);
            xmlq.setAttributes(gradientFill, {
                type: fill.gradientType === "path" ? "path" : undefined,
                left: fill.left,
                right: fill.right,
                top: fill.top,
                bottom: fill.bottom,
                degree: fill.angle
            });

            _.forEach(fill.stops, (fillStop, i) => {
                const stop = {
                    name: 'stop',
                    attributes: { position: fillStop.position },
                    children: []
                };
                gradientFill.children.push(stop);
                this._setColor(stop, 'color', fillStop.color);
            });

            return;
        }

        // Solid fill (really a pattern fill with a solid pattern type).
        if (!_.isObject(fill)) fill = { type: "solid", color: fill };
        else if (fill.hasOwnProperty('rgb') || fill.hasOwnProperty("theme")) fill = { color: fill };

        const patternFill = {
            name: 'patternFill',
            attributes: { patternType: 'solid' }
        };
        this._fillNode.children.push(patternFill);
        this._setColor(patternFill, "fgColor", fill.color);
    }

    _getBorder() {
        const result = {};
        ["left", "right", "top", "bottom", "diagonal"].forEach(side => {
            const sideNode = xmlq.findChild(this._borderNode, side);
            const sideResult = {};

            const style = xmlq.getChildAttribute(this._borderNode, side, 'style');
            if (style) sideResult.style = style;
            const color =this._getColor(sideNode, 'color');
            if (color) sideResult.color = color;

            if (side === "diagonal") {
                const up = this._borderNode.attributes.diagonalUp;
                const down = this._borderNode.attributes.diagonalDown;
                let direction;
                if (up && down) direction = "both";
                else if (up) direction = "up";
                else if (down) direction = "down";
                if (direction) sideResult.direction = direction;
            }

            if (!_.isEmpty(sideResult)) result[side] = sideResult;
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
                xmlq.setChildAttributes(this._borderNode, side, { style: setting.style });
            }

            if (setting.hasOwnProperty("color")) {
                const sideNode = xmlq.findChild(this._borderNode, side);
                this._setColor(sideNode, "color", setting.color);
            }

            if (side === "diagonal") {
                xmlq.setAttributes(this._borderNode, {
                    diagonalUp: setting.direction === "up" || setting.direction === "both" ? 1 : null,
                    diagonalDown: setting.direction === "down" || setting.direction === "both" ? 1 : null
                });
            }
        });
    }

    _get_border() {
        return this._getBorder();
    }

    _set_border(settings) {
        if (_.isObject(settings) && !jq.has(settings, "style") && !jq.has(settings, "color")) {
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

    _get_numberFormat() {
        debug("_get_numberFormat(%o)", arguments);
        const numFmtId = this._xfNode.attributes.numFmtId || 0;
        return this._styleSheet.getNumberFormatCode(numFmtId);
    }

    _set_numberFormat(formatCode) {
        debug("_set_numberFormat(%o)", arguments);
        this._xfNode.attributes.numFmtId = this._styleSheet.getNumberFormatId(formatCode);
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
