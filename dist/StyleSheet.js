"use strict";
const _ = require("lodash");
const xmlq = require("./xmlq");
const Style = require("./Style");
/**
 * Standard number format codes
 * Taken from http://polymathprogrammer.com/2011/02/15/built-in-styles-for-excel-open-xml/
 * @private
 */
const STANDARD_CODES = {
    0: 'General',
    1: '0',
    2: '0.00',
    3: '#,##0',
    4: '#,##0.00',
    9: '0%',
    10: '0.00%',
    11: '0.00E+00',
    12: '# ?/?',
    13: '# ??/??',
    14: 'mm-dd-yy',
    15: 'd-mmm-yy',
    16: 'd-mmm',
    17: 'mmm-yy',
    18: 'h:mm AM/PM',
    19: 'h:mm:ss AM/PM',
    20: 'h:mm',
    21: 'h:mm:ss',
    22: 'm/d/yy h:mm',
    37: '#,##0 ;(#,##0)',
    38: '#,##0 ;[Red](#,##0)',
    39: '#,##0.00;(#,##0.00)',
    40: '#,##0.00;[Red](#,##0.00)',
    45: 'mm:ss',
    46: '[h]:mm:ss',
    47: 'mmss.0',
    48: '##0.0E+0',
    49: '@'
};
/**
 * The starting ID for custom number formats. The first 163 indexes are reserved.
 * @private
 */
const STARTING_CUSTOM_NUMBER_FORMAT_ID = 164;
/**
 * A style sheet.
 * @ignore
 */
class StyleSheet {
    /**
     * Creates an instance of _StyleSheet.
     * @param {string} node - The style sheet node
     */
    constructor(node) {
        this._init(node);
        this._cacheNumberFormats();
    }
    /**
     * Create a style.
     * @param {number} [sourceId] - The source style ID to copy, if provided.
     * @returns {Style} The style.
     */
    createStyle(sourceId) {
        let fontNode, fillNode, borderNode, xfNode;
        if (sourceId >= 0) {
            const sourceXfNode = this._cellXfsNode.children[sourceId];
            xfNode = _.cloneDeep(sourceXfNode);
            if (sourceXfNode.attributes.applyFont) {
                const fontId = sourceXfNode.attributes.fontId;
                fontNode = _.cloneDeep(this._fontsNode.children[fontId]);
            }
            if (sourceXfNode.attributes.applyFill) {
                const fillId = sourceXfNode.attributes.fillId;
                fillNode = _.cloneDeep(this._fillsNode.children[fillId]);
            }
            if (sourceXfNode.attributes.applyBorder) {
                const borderId = sourceXfNode.attributes.borderId;
                borderNode = _.cloneDeep(this._bordersNode.children[borderId]);
            }
        }
        if (!fontNode)
            fontNode = { name: "font", attributes: {}, children: [] };
        this._fontsNode.children.push(fontNode);
        if (!fillNode)
            fillNode = { name: "fill", attributes: {}, children: [] };
        this._fillsNode.children.push(fillNode);
        // The border sides must be in order
        if (!borderNode)
            borderNode = { name: "border", attributes: {}, children: [] };
        borderNode.children = [
            xmlq.findChild(borderNode, "left") || { name: "left", attributes: {}, children: [] },
            xmlq.findChild(borderNode, "right") || { name: "right", attributes: {}, children: [] },
            xmlq.findChild(borderNode, "top") || { name: "top", attributes: {}, children: [] },
            xmlq.findChild(borderNode, "bottom") || { name: "bottom", attributes: {}, children: [] },
            xmlq.findChild(borderNode, "diagonal") || { name: "diagonal", attributes: {}, children: [] }
        ];
        this._bordersNode.children.push(borderNode);
        if (!xfNode)
            xfNode = { name: "xf", attributes: {}, children: [] };
        _.assign(xfNode.attributes, {
            fontId: this._fontsNode.children.length - 1,
            fillId: this._fillsNode.children.length - 1,
            borderId: this._bordersNode.children.length - 1,
            applyFont: 1,
            applyFill: 1,
            applyBorder: 1
        });
        this._cellXfsNode.children.push(xfNode);
        return new Style(this, this._cellXfsNode.children.length - 1, xfNode, fontNode, fillNode, borderNode);
    }
    /**
     * Get the number format code for a given ID.
     * @param {number} id - The number format ID.
     * @returns {string} The format code.
     */
    getNumberFormatCode(id) {
        return this._numberFormatCodesById[id];
    }
    /**
     * Get the nuumber format ID for a given code.
     * @param {string} code - The format code.
     * @returns {number} The number format ID.
     */
    getNumberFormatId(code) {
        let id = this._numberFormatIdsByCode[code];
        if (id === undefined) {
            id = this._nextNumFormatId++;
            this._numberFormatCodesById[id] = code;
            this._numberFormatIdsByCode[code] = id;
            this._numFmtsNode.children.push({
                name: 'numFmt',
                attributes: {
                    numFmtId: id,
                    formatCode: code
                }
            });
        }
        return id;
    }
    /**
     * Convert the style sheet to an XML object.
     * @returns {{}} The XML form.
     * @ignore
     */
    toXml() {
        return this._node;
    }
    /**
     * Cache the number format codes
     * @returns {undefined}
     * @private
     */
    _cacheNumberFormats() {
        // Load the standard number format codes into the caches.
        this._numberFormatCodesById = {};
        this._numberFormatIdsByCode = {};
        for (const id in STANDARD_CODES) {
            if (!STANDARD_CODES.hasOwnProperty(id))
                continue;
            const code = STANDARD_CODES[id];
            this._numberFormatCodesById[id] = code;
            this._numberFormatIdsByCode[code] = parseInt(id);
        }
        // Set the next number format code.
        this._nextNumFormatId = STARTING_CUSTOM_NUMBER_FORMAT_ID;
        // If there are custom number formats, cache them all and update the next num as needed.
        this._numFmtsNode.children.forEach(node => {
            const id = node.attributes.numFmtId;
            const code = node.attributes.formatCode;
            this._numberFormatCodesById[id] = code;
            this._numberFormatIdsByCode[code] = id;
            if (id >= this._nextNumFormatId)
                this._nextNumFormatId = id + 1;
        });
    }
    /**
     * Initialize the style sheet node.
     * @param {{}} [node] - The node
     * @returns {undefined}
     * @private
     */
    _init(node) {
        this._node = node;
        // Cache the refs to the collections.
        this._numFmtsNode = xmlq.findChild(this._node, "numFmts");
        this._fontsNode = xmlq.findChild(this._node, "fonts");
        this._fillsNode = xmlq.findChild(this._node, "fills");
        this._bordersNode = xmlq.findChild(this._node, "borders");
        this._cellXfsNode = xmlq.findChild(this._node, "cellXfs");
        if (!this._numFmtsNode) {
            this._numFmtsNode = {
                name: "numFmts",
                attributes: {},
                children: []
            };
            // Number formats need to be before the others.
            xmlq.insertBefore(this._node, this._numFmtsNode, this._fontsNode);
        }
        // Remove the optional counts so we don't have to keep them up to date.
        delete this._numFmtsNode.attributes.count;
        delete this._fontsNode.attributes.count;
        delete this._fillsNode.attributes.count;
        delete this._bordersNode.attributes.count;
        delete this._cellXfsNode.attributes.count;
    }
}
module.exports = StyleSheet;
//# sourceMappingURL=StyleSheet.js.map