"use strict";

/* eslint camelcase:off */

const ArgHandler = require("./ArgHandler");
const _ = require("lodash");
const xmlq = require("./xmlq");
const colorIndexes = require("./colorIndexes");

/**
 * Rich text.
 * @ignore
 */
class RichText {
    /**
     * Creates a new instance of RichText.
     * @constructor
     * @param {string} value - text value
     * @param {object|undefined|null} styles - multiple styles
     * @param {Cell|undefined|null} cell - the cell that display the rich text
     */
    constructor(value, styles, cell) {
        this._cell = cell;
        if (value.name === 'r') {
            this._node = value;
            this._fontNode = xmlq.findChild(this._node, 'rPr');
            if (!this._fontNode) {
                this._fontNode = { name: 'rPr', attributes: {}, children: [] };
                this._node.children.unshift(this._fontNode);
            }
            this._valueNode = xmlq.findChild(this._node, 't');
        } else {
            this._node = {
                name: 'r',
                attributes: {},
                children: [
                    { name: 'rPr', attributes: {}, children: [] },
                    { name: 't', attributes: {}, children: [] }
                ]
            };
            this._fontNode = xmlq.findChild(this._node, 'rPr');
            this._valueNode = xmlq.findChild(this._node, 't');
            this.value(value);
        }
    }

    value() {
        return new ArgHandler("_RichText.value")
            .case(() => {
                return this._valueNode.children[0];
            })
            .case('string', value => {
                value = value.replace(/(?:\r\n|\r|\n)/g, '\r\n');
                const hasLineSeparator = value.indexOf('\r\n') !== -1;
                this._valueNode.children[0] = value;

                if (hasLineSeparator) {
                    // set wrapText = true if it contains line separator, excel will only display new lines if it sets.
                    if (this._cell) {
                        this._cell.style('wrapText', true);
                    }
                    xmlq.setAttributes(this._valueNode, { 'xml:space': 'preserve' });
                }
            })
            .handle(arguments);
    }

    toXml() {
        return this._node;
    }

    /**
     * Gets or sets a style.
     * @param {string} name - The style name.
     * @param {*} [value] - The value to set.
     * @returns {*|RichText} The value if getting or the style if setting.
     */
    style() {
        return new ArgHandler("_RichText.style")
            .case('string', name => {
                // Get single value
                const getterName = `_get_${name}`;
                if (!this[getterName]) throw new Error(`_RichText.style: '${name}' is not a valid style`);
                return this[getterName]();
            })
            .case('array', names => {
                // Get list of values
                const values = {};
                names.forEach(name => {
                    values[name] = this.style(name);
                });
                return values;
            })
            .case(['string', '*'], (name, value) => {
                // Set a single value
                const setterName = `_set_${name}`;
                if (!this[setterName]) throw new Error(`_RichText.style: '${name}' is not a valid style`);
                return this[setterName](value);
            })
            .case('object', nameValues => {
                // Object of key value pairs to set
                for (const name in nameValues) {
                    if (!nameValues.hasOwnProperty(name)) continue;
                    const value = nameValues[name];
                    this.style(name, value);
                }

                return this;
            })
            .case('Style', style => {
                this._style = style;
                this._styleId = style.id();

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
        return xmlq.getChildAttribute(this._fontNode, 'rFont', "val");
    }

    _set_fontFamily(family) {
        xmlq.setChildAttributes(this._fontNode, 'rFont', { val: family });
        xmlq.removeChildIfEmpty(this._fontNode, 'rFont');
    }

    _get_fontGenericFamily() {
        return xmlq.getChildAttribute(this._fontNode, 'family', "val");
    }

    _set_fontGenericFamily(genericFamily) {
        xmlq.setChildAttributes(this._fontNode, 'family', { val: genericFamily });
        xmlq.removeChildIfEmpty(this._fontNode, 'family');
    }

    _get_fontColor() {
        return this._getColor(this._fontNode, "color");
    }

    _set_fontColor(color) {
        this._setColor(this._fontNode, "color", color);
    }

    _get_fontScheme() {
        // can be 'minor', 'major', 'none'
        return xmlq.getChildAttribute(this._fontNode, 'scheme', "val");
    }

    _set_fontScheme(scheme) {
        xmlq.setChildAttributes(this._fontNode, 'scheme', { val: scheme });
        xmlq.removeChildIfEmpty(this._fontNode, 'scheme');
    }
}

// IE doesn't support function names so explicitly set it.
if (!RichText.name) RichText.name = "RichText";

module.exports = RichText;
