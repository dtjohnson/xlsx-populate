"use strict";

/* eslint camelcase:off */

const ArgHandler = require("./ArgHandler");
const _ = require("lodash");
const xmlq = require("./xmlq");
const colorIndexes = require("./colorIndexes");

/**
 * A Rich text fragment.
 */
class RichTextFragment {
    /**
     * Creates a new instance of RichTextFragment.
     * @constructor
     * @param {string|Object} value - Text value or XML node
     * @param {object|undefined|null} [styles] - Multiple styles.
     * @param {RichText} richText - The rich text instance where this fragment belongs to.
     */
    constructor(value, styles, richText) {
        this._richText = richText;
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
            if (styles) {
                this.style(styles);
            }
        }
    }

    /**
     * Gets the value of this part of rich text
     * @return {string} text
     *//**
     * Sets the value of this part of rich text
     * @param {string} text - the text to set
     * @return {RichTextFragment} - RichTextFragment
     */
    value() {
        return new ArgHandler("_RichText.value")
            .case(() => {
                return this._valueNode.children[0];
            })
            .case('string', value => {
                value = value.replace(/(?:\r\n|\r|\n)/g, '\r\n');
                const hasLineSeparator = value.indexOf('\r\n') !== -1;
                this._valueNode.children[0] = value;
                if (value.charAt(0) === ' ') xmlq.setAttributes(this._valueNode, { 'xml:space': 'preserve' });

                if (this._richText) this._richText.removeUnsupportedNodes();
                if (hasLineSeparator) {
                    // set wrapText = true if it contains line separator, excel will only display new lines if it sets.
                    if (this._richText.cell) {
                        this._richText.cell.style('wrapText', true);
                    }
                    xmlq.setAttributes(this._valueNode, { 'xml:space': 'preserve' });
                }
                return this;
            })
            .handle(arguments);
    }

    /**
     * Convert the rich text to an XML object.
     * @returns {{}} The XML form.
     * @ignore
     */
    toXml() {
        return this._node;
    }

    /**
     * Gets an individual style.
     * @param {string} name - The name of the style.
     * @returns {*} The style.
     *//**
     * Gets multiple styles.
     * @param {Array.<string>} names - The names of the style.
     * @returns {object.<string, *>} Object whose keys are the style names and values are the styles.
     *//**
     * Sets an individual style.
     * @param {string} name - The name of the style.
     * @param {*} value - The value to set.
     * @returns {RichTextFragment} This RichTextFragment.
     *//**
     * Sets multiple styles.
     * @param {object.<string, *>} styles - Object whose keys are the style names and values are the styles to set.
     * @returns {RichTextFragment} This RichTextFragment.
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

    /**
     * @param {number} genericFamily - 1: Serif, 2: Sans Serif, 3: Monospace,
     * @private
     * @return {undefined}
     */
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

    /**
     * @param {string} scheme - 'minor'|'major'|'none'
     * @private
     * @return {undefined}
     */
    _set_fontScheme(scheme) {
        xmlq.setChildAttributes(this._fontNode, 'scheme', { val: scheme });
        xmlq.removeChildIfEmpty(this._fontNode, 'scheme');
    }
}

// IE doesn't support function names so explicitly set it.
if (!RichTextFragment.name) RichTextFragment.name = "RichTextFragment";

module.exports = RichTextFragment;
