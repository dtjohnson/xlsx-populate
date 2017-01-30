"use strict";

// TODO: Tests

// TODO: Support sheet, row, and column styles.

const debug = require("debug")("_StyleSheet");
const utils = require("./utils");
const _Style = require("./_Style");
const DOMParser = require('xmldom').DOMParser;
const parser = new DOMParser();

// Taken from http://polymathprogrammer.com/2011/02/15/built-in-styles-for-excel-open-xml/
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
 * A style sheet.
 */
class _StyleSheet {
    /**
     * Creates an instance of _StyleSheet.
     * @param {string} text - The style sheet XML text.
     */
    constructor(text) {
        debug("constructor(_)");

        // Parse the XML.
        this._xml = parser.parseFromString(text);

        // Load the standard number format codes into the caches.
        this._numberFormatCodesById = {};
        this._numberFormatIdsByCode = {};
        for (const id in STANDARD_CODES) {
            if (!STANDARD_CODES.hasOwnProperty(id)) continue;
            const code = STANDARD_CODES[id];
            this._numberFormatCodesById[id] = code;
            this._numberFormatIdsByCode[code] = parseInt(id);
        }

        // Set the next number format code to 164. The first 163 indexes are reserved.
        this._nextNumFormatId = 164;

        // If there are custom number formats, cache them all and update the next num as needed.
        this._numFmtsNode = this._xml.documentElement.getElementsByTagName("numFmts")[0];
        if (this._numFmtsNode) {
            this._numFmtsNode.removeAttribute("count");
            utils.forEachChildElement(this._numFmtsNode, node => {
                const id = parseInt(node.getAttribute("numFmtId"));
                const code = node.getAttribute("formatCode");
                this._numberFormatCodesById[id] = code;
                this._numberFormatIdsByCode[code] = id;
                if (id >= this._nextNumFormatId) this._nextNumFormatId = id + 1;
            });
        } else {
            // Create the numFmts node if it doesn't exist. It must be first in the style sheet.
            this._numFmtsNode = this._xml.createElement("numFmts");
            this._xml.documentElement.insertBefore(this._numFmtsNode, this._xml.documentElement.firstChild);
        }

        // Grab the other style collection nodes and remove the counts. They aren't mandatory.
        this._fontsNode = utils.findFirstElement(this._xml.documentElement, "fonts");
        this._fontsNode.setAttribute("count", utils.findElements(this._fontsNode).length);

        this._fillsNode = utils.findFirstElement(this._xml.documentElement, "fills");
        this._fillsNode.setAttribute("count", utils.findElements(this._fillsNode).length);

        this._bordersNode = utils.findFirstElement(this._xml.documentElement, "borders");
        this._bordersNode.setAttribute("count", utils.findElements(this._bordersNode).length);

        this._cellXfsNode = utils.findFirstElement(this._xml.documentElement, "cellXfs");
        this._cellXfsNode.setAttribute("count", utils.findElements(this._cellXfsNode).length);
    }

    /**
     * Create a style.
     * @param {number} [sourceId] - The source style ID to copy, if provided.
     * @returns {_Style} The style.
     */
    createStyle(sourceId) {
        debug("createStyle(%o)", arguments);

        let fontNode, fillNode, borderNode, xfNode;
        if (sourceId !== undefined) {
            debug("Cloning source xf node");
            const sourceXfNode = utils.findNthElement(this._cellXfsNode, sourceId);
            xfNode = sourceXfNode.cloneNode(true);

            if (sourceXfNode.getAttribute("applyFont") === "1") {
                const fontId = parseInt(sourceXfNode.getAttribute("fontId"));
                debug("Cloning source font node: %s", fontId);
                fontNode = this._fontsNode.childNodes[fontId].cloneNode(true);
            }

            if (sourceXfNode.getAttribute("applyFill") === "1") {
                const fillId = parseInt(sourceXfNode.getAttribute("fillId"));
                debug("Cloning source fill node: %s", fillId);
                fillNode = this._fillsNode.childNodes[fillId].cloneNode(true);
            }

            if (sourceXfNode.getAttribute("applyBorder") === "1") {
                const borderId = parseInt(sourceXfNode.getAttribute("borderId"));
                debug("Cloning source font node: %s", borderId);
                borderNode = this._bordersNode.childNodes[borderId].cloneNode(true);
            }
        }

        if (!fontNode) fontNode = this._xml.createElement("font");
        this._fontsNode.appendChild(fontNode);

        if (!fillNode) fillNode = this._xml.createElement("fill");
        this._fillsNode.appendChild(fillNode);

        if (!borderNode) borderNode = this._xml.createElement("border");
        this._bordersNode.appendChild(borderNode);

        const fontCount = parseInt(this._fontsNode.getAttribute("count")) + 1;
        this._fontsNode.setAttribute("count", fontCount);

        const fillCount = parseInt(this._fillsNode.getAttribute("count")) + 1;
        this._fillsNode.setAttribute("count", fillCount);

        const borderCount = parseInt(this._bordersNode.getAttribute("count")) + 1;
        this._bordersNode.setAttribute("count", borderCount);

        const cellXfCount = parseInt(this._cellXfsNode.getAttribute("count")) + 1;
        this._cellXfsNode.setAttribute("count", cellXfCount);

        if (!xfNode) xfNode = this._xml.createElement("xf");
        xfNode.setAttribute("fontId", fontCount - 1);
        xfNode.setAttribute("fillId", fillCount - 1);
        xfNode.setAttribute("borderId", borderCount - 1);
        xfNode.setAttribute("applyFont", 1);
        xfNode.setAttribute("applyFill", 1);
        xfNode.setAttribute("applyBorder", 1);
        this._cellXfsNode.appendChild(xfNode);

        return new _Style(this, cellXfCount - 1, xfNode, fontNode, fillNode, borderNode);
    }

    /**
     * Get the number format code for a given ID.
     * @param {number} id - The number format ID.
     * @returns {string} The format code.
     */
    getNumberFormatCode(id) {
        debug("getNumberFormatCode(%o)", arguments);
        return this._numberFormatCodesById[id];
    }

    /**
     * Get the nuumber format ID for a given code.
     * @param {string} code - The format code.
     * @returns {number} The number format ID.
     */
    getNumberFormatId(code) {
        debug("getNumberFormatId(%o)", arguments);
        let id = this._numberFormatIdsByCode[code];
        if (id === undefined) {
            id = this._nextNumFormatId++;
            this._numberFormatCodesById[id] = code;
            this._numberFormatIdsByCode[code] = id;

            const numFmtNode = this._xml.createElement("numFmt");
            numFmtNode.setAttribute("numFmtId", id);
            numFmtNode.setAttribute("formatCode", code);
            this._numFmtsNode.appendChild(numFmtNode);
        }

        return id;
    }

    /**
     * Convert the style sheet to an XML string.
     * @returns {string} The XML string.
     */
    toString() {
        debug("toString(%o)", arguments);
        return this._xml.toString();
    }
}

module.exports = _StyleSheet;

/*
xl/styles.xml

<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main">
    <numFmts count="1">
        <numFmt numFmtId="164" formatCode="#,##0_);[Red]\(#,##0\)\)"/>
    </numFmts>
    <fonts count="1" x14ac:knownFonts="1">
        <font>
            <sz val="11"/>
            <color theme="1"/>
            <name val="Calibri"/>
            <family val="2"/>
            <scheme val="minor"/>
        </font>
    </fonts>
    <fills count="11">
        <fill>
            <patternFill patternType="none"/>
        </fill>
        <fill>
            <patternFill patternType="gray125"/>
        </fill>
        <fill>
            <patternFill patternType="solid">
                <fgColor rgb="FFC00000"/>
                <bgColor indexed="64"/>
            </patternFill>
        </fill>
        <fill>
            <patternFill patternType="lightDown">
                <fgColor theme="4"/>
                <bgColor rgb="FFC00000"/>
            </patternFill>
        </fill>
        <fill>
            <gradientFill degree="90">
                <stop position="0">
                    <color theme="0"/>
                </stop>
                <stop position="1">
                    <color theme="4"/>
                </stop>
            </gradientFill>
        </fill>
        <fill>
            <gradientFill>
                <stop position="0">
                    <color theme="0"/>
                </stop>
                <stop position="1">
                    <color theme="4"/>
                </stop>
            </gradientFill>
        </fill>
        <fill>
            <gradientFill degree="45">
                <stop position="0">
                    <color theme="0"/>
                </stop>
                <stop position="1">
                    <color theme="4"/>
                </stop>
            </gradientFill>
        </fill>
        <fill>
            <gradientFill degree="135">
                <stop position="0">
                    <color theme="0"/>
                </stop>
                <stop position="1">
                    <color theme="4"/>
                </stop>
            </gradientFill>
        </fill>
        <fill>
            <gradientFill type="path">
                <stop position="0">
                    <color theme="0"/>
                </stop>
                <stop position="1">
                    <color theme="4"/>
                </stop>
            </gradientFill>
        </fill>
        <fill>
            <gradientFill type="path" left="0.5" right="0.5" top="0.5" bottom="0.5">
                <stop position="0">
                    <color theme="0"/>
                </stop>
                <stop position="1">
                    <color theme="4"/>
                </stop>
            </gradientFill>
        </fill>
        <fill>
            <gradientFill degree="270">
                <stop position="0">
                    <color theme="0"/>
                </stop>
                <stop position="1">
                    <color theme="4"/>
                </stop>
            </gradientFill>
        </fill>
    </fills>
    <borders count="10">
        <border>
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal/>
        </border>
        <border diagonalDown="1">
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal style="hair">
                <color auto="1"/>
            </diagonal>
        </border>
        <border diagonalDown="1">
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal style="dotted">
                <color auto="1"/>
            </diagonal>
        </border>
        <border diagonalDown="1">
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal style="dashDotDot">
                <color auto="1"/>
            </diagonal>
        </border>
        <border diagonalDown="1">
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal style="dashDot">
                <color auto="1"/>
            </diagonal>
        </border>
        <border diagonalDown="1">
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal style="dashed">
                <color auto="1"/>
            </diagonal>
        </border>
        <border diagonalUp="1">
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal style="mediumDashDotDot">
                <color auto="1"/>
            </diagonal>
        </border>
        <border diagonalUp="1">
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal style="slantDashDot">
                <color auto="1"/>
            </diagonal>
        </border>
        <border diagonalUp="1">
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal style="mediumDashDot">
                <color auto="1"/>
            </diagonal>
        </border>
        <border diagonalUp="1">
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal style="mediumDashed">
                <color auto="1"/>
            </diagonal>
        </border>
    </borders>
    <cellStyleXfs count="1">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    </cellStyleXfs>
    <cellXfs count="19">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="2" xfId="0" applyBorder="1"/>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="3" xfId="0" applyBorder="1"/>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="4" xfId="0" applyBorder="1"/>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="5" xfId="0" applyBorder="1"/>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="6" xfId="0" applyBorder="1"/>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="7" xfId="0" applyBorder="1"/>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="8" xfId="0" applyBorder="1"/>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="9" xfId="0" applyBorder="1"/>
        <xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1" applyBorder="1"/>
        <xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/>
        <xf numFmtId="0" fontId="0" fillId="4" borderId="0" xfId="0" applyFill="1"/>
        <xf numFmtId="0" fontId="0" fillId="5" borderId="0" xfId="0" applyFill="1"/>
        <xf numFmtId="0" fontId="0" fillId="6" borderId="0" xfId="0" applyFill="1"/>
        <xf numFmtId="0" fontId="0" fillId="7" borderId="0" xfId="0" applyFill="1"/>
        <xf numFmtId="0" fontId="0" fillId="8" borderId="0" xfId="0" applyFill="1"/>
        <xf numFmtId="0" fontId="0" fillId="9" borderId="0" xfId="0" applyFill="1"/>
        <xf numFmtId="0" fontId="0" fillId="10" borderId="0" xfId="0" applyFill="1"/>
    </cellXfs>
    <cellStyles count="1">
        <cellStyle name="Normal" xfId="0" builtinId="0"/>
    </cellStyles>
    <dxfs count="0"/>
    <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
    <extLst>
        <ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
            <x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/>
        </ext>
        <ext uri="{9260A510-F301-46a8-8635-F512D64BE5F5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
            <x15:timelineStyles defaultTimelineStyle="TimeSlicerStyleLight1"/>
        </ext>
    </extLst>
</styleSheet>
*/
