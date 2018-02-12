"use strict";

/* eslint camelcase:off */

const ArgHandler = require("./ArgHandler");
const _ = require("lodash");
const xmlq = require("./xmlq");
const colorIndexes = require("./colorIndexes");

/**
 * A style.
 * @ignore
 */
class Style {
    /**
     * Creates a new instance of _Style.
     * @constructor
     * @param {StyleSheet} styleSheet - The styleSheet.
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
     * Gets the style ID.
     * @returns {number} The ID.
     */
    id() {
        return this._id;
    }

    /**
     * Gets or sets a style.
     * @param {string} name - The style name.
     * @param {*} [value] - The value to set.
     * @returns {*|Style} The value if getting or the style if setting.
     */
    style() {
        return new ArgHandler("_Style.style")
            .case('string', name => {
                const getterName = `_get_${name}`;
                if (!this[getterName]) throw new Error(`_Style.style: '${name}' is not a valid style`);
                return this[getterName]();
            })
            .case(['string', '*'], (name, value) => {
                const setterName = `_set_${name}`;
                if (!this[setterName]) throw new Error(`_Style.style: '${name}' is not a valid style`);
                this[setterName](value);
                return this;
            })
            .handle(arguments);
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
        return xmlq.hasChild(this._fontNode, 'b');
    }

    _set_bold(bold) {
        if (bold) xmlq.appendChildIfNotFound(this._fontNode, "b");
        else xmlq.removeChild(this._fontNode, 'b');
    }

    _get_italic() {
        return xmlq.hasChild(this._fontNode, 'i');
    }

    _set_italic(italic) {
        if (italic) xmlq.appendChildIfNotFound(this._fontNode, "i");
        else xmlq.removeChild(this._fontNode, 'i');
    }

    _get_underline() {
        const uNode = xmlq.findChild(this._fontNode, 'u');
        return uNode ? uNode.attributes.val || true : false;
    }

    _set_underline(underline) {
        if (underline) {
            const uNode = xmlq.appendChildIfNotFound(this._fontNode, "u");
            const val = typeof underline === 'string' ? underline : null;
            xmlq.setAttributes(uNode, { val });
        } else {
            xmlq.removeChild(this._fontNode, 'u');
        }
    }

    _get_strikethrough() {
        return xmlq.hasChild(this._fontNode, 'strike');
    }

    _set_strikethrough(strikethrough) {
        if (strikethrough) xmlq.appendChildIfNotFound(this._fontNode, "strike");
        else xmlq.removeChild(this._fontNode, 'strike');
    }

    _getFontVerticalAlignment() {
        return xmlq.getChildAttribute(this._fontNode, 'vertAlign', "val");
    }

    _setFontVerticalAlignment(alignment) {
        xmlq.setChildAttributes(this._fontNode, 'vertAlign', { val: alignment });
        xmlq.removeChildIfEmpty(this._fontNode, 'vertAlign');
    }

    _get_subscript() {
        return this._getFontVerticalAlignment() === "subscript";
    }

    _set_subscript(subscript) {
        this._setFontVerticalAlignment(subscript ? "subscript" : null);
    }

    _get_superscript() {
        return this._getFontVerticalAlignment() === "superscript";
    }

    _set_superscript(superscript) {
        this._setFontVerticalAlignment(superscript ? "superscript" : null);
    }

    _get_fontSize() {
        return xmlq.getChildAttribute(this._fontNode, 'sz', "val");
    }

    _set_fontSize(size) {
        xmlq.setChildAttributes(this._fontNode, 'sz', { val: size });
        xmlq.removeChildIfEmpty(this._fontNode, 'sz');
    }

    _get_fontFamily() {
        return xmlq.getChildAttribute(this._fontNode, 'name', "val");
    }

    _set_fontFamily(family) {
        xmlq.setChildAttributes(this._fontNode, 'name', { val: family });
        xmlq.removeChildIfEmpty(this._fontNode, 'name');
    }

    _get_fontColor() {
        return this._getColor(this._fontNode, "color");
    }

    _set_fontColor(color) {
        this._setColor(this._fontNode, "color", color);
    }

    _get_horizontalAlignment() {
        return xmlq.getChildAttribute(this._xfNode, 'alignment', "horizontal");
    }

    _set_horizontalAlignment(alignment) {
        xmlq.setChildAttributes(this._xfNode, 'alignment', { horizontal: alignment });
        xmlq.removeChildIfEmpty(this._xfNode, 'alignment');
    }

    _get_justifyLastLine() {
        return xmlq.getChildAttribute(this._xfNode, 'alignment', "justifyLastLine") === 1;
    }

    _set_justifyLastLine(justifyLastLine) {
        xmlq.setChildAttributes(this._xfNode, 'alignment', { justifyLastLine: justifyLastLine ? 1 : null });
        xmlq.removeChildIfEmpty(this._xfNode, 'alignment');
    }

    _get_indent() {
        return xmlq.getChildAttribute(this._xfNode, 'alignment', "indent");
    }

    _set_indent(indent) {
        xmlq.setChildAttributes(this._xfNode, 'alignment', { indent });
        xmlq.removeChildIfEmpty(this._xfNode, 'alignment');
    }

    _get_verticalAlignment() {
        return xmlq.getChildAttribute(this._xfNode, 'alignment', "vertical");
    }

    _set_verticalAlignment(alignment) {
        xmlq.setChildAttributes(this._xfNode, 'alignment', { vertical: alignment });
        xmlq.removeChildIfEmpty(this._xfNode, 'alignment');
    }

    _get_wrapText() {
        return xmlq.getChildAttribute(this._xfNode, 'alignment', "wrapText") === 1;
    }

    _set_wrapText(wrapText) {
        xmlq.setChildAttributes(this._xfNode, 'alignment', { wrapText: wrapText ? 1 : null });
        xmlq.removeChildIfEmpty(this._xfNode, 'alignment');
    }

    _get_shrinkToFit() {
        return xmlq.getChildAttribute(this._xfNode, 'alignment', "shrinkToFit") === 1;
    }

    _set_shrinkToFit(shrinkToFit) {
        xmlq.setChildAttributes(this._xfNode, 'alignment', { shrinkToFit: shrinkToFit ? 1 : null });
        xmlq.removeChildIfEmpty(this._xfNode, 'alignment');
    }

    _get_textDirection() {
        const readingOrder = xmlq.getChildAttribute(this._xfNode, 'alignment', "readingOrder");
        if (readingOrder === 1) return "left-to-right";
        if (readingOrder === 2) return "right-to-left";
        return readingOrder;
    }

    _set_textDirection(textDirection) {
        let readingOrder;
        if (textDirection === "left-to-right") readingOrder = 1;
        else if (textDirection === "right-to-left") readingOrder = 2;
        xmlq.setChildAttributes(this._xfNode, 'alignment', { readingOrder });
        xmlq.removeChildIfEmpty(this._xfNode, 'alignment');
    }

    _getTextRotation() {
        return xmlq.getChildAttribute(this._xfNode, 'alignment', "textRotation");
    }

    _setTextRotation(textRotation) {
        xmlq.setChildAttributes(this._xfNode, 'alignment', { textRotation });
        xmlq.removeChildIfEmpty(this._xfNode, 'alignment');
    }

    _get_textRotation() {
        let textRotation = this._getTextRotation();

        // Negative angles in Excel correspond to values > 90 in OOXML.
        if (textRotation > 90) textRotation = 90 - textRotation;
        return textRotation;
    }

    _set_textRotation(textRotation) {
        // Negative angles in Excel correspond to values > 90 in OOXML.
        if (textRotation < 0) textRotation = 90 - textRotation;
        this._setTextRotation(textRotation);
    }

    _get_angleTextCounterclockwise() {
        return this._getTextRotation() === 45;
    }

    _set_angleTextCounterclockwise(value) {
        this._setTextRotation(value ? 45 : null);
    }

    _get_angleTextClockwise() {
        return this._getTextRotation() === 135;
    }

    _set_angleTextClockwise(value) {
        this._setTextRotation(value ? 135 : null);
    }

    _get_rotateTextUp() {
        return this._getTextRotation() === 90;
    }

    _set_rotateTextUp(value) {
        this._setTextRotation(value ? 90 : null);
    }

    _get_rotateTextDown() {
        return this._getTextRotation() === 180;
    }

    _set_rotateTextDown(value) {
        this._setTextRotation(value ? 180 : null);
    }

    _get_verticalText() {
        return this._getTextRotation() === 255;
    }

    _set_verticalText(value) {
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
                fill.angle = gradientFillNode.attributes.degree;
            } else {
                fill.left = gradientFillNode.attributes.left;
                fill.right = gradientFillNode.attributes.right;
                fill.top = gradientFillNode.attributes.top;
                fill.bottom = gradientFillNode.attributes.bottom;
            }

            return fill;
        }
    }

    _set_fill(fill) {
        this._fillNode.children = [];

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
            this._setColor(patternFill, "fgColor", fill.foreground);
            this._setColor(patternFill, "bgColor", fill.background);
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
            const color = this._getColor(sideNode, 'color');
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
        if (_.isObject(settings) && !settings.hasOwnProperty("style") && !settings.hasOwnProperty("color")) {
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
        const border = this._getBorder().diagonal;
        return border && border.direction;
    }

    _set_diagonalBorderDirection(direction) {
        this._setBorder({ diagonal: { direction } });
    }

    _get_numberFormat() {
        const numFmtId = this._xfNode.attributes.numFmtId || 0;
        return this._styleSheet.getNumberFormatCode(numFmtId);
    }

    _set_numberFormat(formatCode) {
        this._xfNode.attributes.numFmtId = this._styleSheet.getNumberFormatId(formatCode);
    }
}

["left", "right", "top", "bottom", "diagonal"].forEach(side => {
    Style.prototype[`_get_${side}Border`] = function () {
        return this._getBorder()[side];
    };

    Style.prototype[`_set_${side}Border`] = function (settings) {
        this._setBorder({ [side]: settings });
    };

    Style.prototype[`_get_${side}BorderColor`] = function () {
        const border = this._getBorder()[side];
        return border && border.color;
    };

    Style.prototype[`_set_${side}BorderColor`] = function (color) {
        this._setBorder({ [side]: { color } });
    };

    Style.prototype[`_get_${side}BorderStyle`] = function () {
        const border = this._getBorder()[side];
        return border && border.style;
    };

    Style.prototype[`_set_${side}BorderStyle`] = function (style) {
        this._setBorder({ [side]: { style } });
    };
});

// IE doesn't support function names so explicitly set it.
if (!Style.name) Style.name = "Style";

module.exports = Style;
