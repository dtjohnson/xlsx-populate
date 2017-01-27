"use strict";

// TODO: Tests
// TODO: JSDoc
// TODO: Debugs
// TODO: Enable number formats

const DOMParser = require('xmldom').DOMParser;
const parser = new DOMParser();
const debug = require("debug")("_StyleSheet");
const utils = require("./utils");
const Style = require("./Style");

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

class _StyleSheet {
    constructor(text) {
        this._xml = parser.parseFromString(text);

        this._numberFormatCodesById = {};
        this._numberFormatIdsByCode = {};
        for (const id in STANDARD_CODES) {
            if (!STANDARD_CODES.hasOwnProperty(id)) continue;
            const code = STANDARD_CODES[id];
            this._numberFormatCodesById[id] = code;
            this._numberFormatIdsByCode[code] = id;
        }

        this._nextNumFormatId = 164;
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
            this._numFmtsNode = this._xml.createElement("numFmts");
            this._xml.documentElement.insertBefore(this._numFmtsNode, this._xml.documentElement.firstChild);
        }

        this._fontsNode = this._xml.documentElement.getElementsByTagName("fonts")[0];
        this._fontsNode.removeAttribute("count");

        this._fillsNode = this._xml.documentElement.getElementsByTagName("fills")[0];
        this._fillsNode.removeAttribute("count");

        this._bordersNode = this._xml.documentElement.getElementsByTagName("borders")[0];
        this._bordersNode.removeAttribute("count");

        this._cellXfsNode = this._xml.documentElement.getElementsByTagName("cellXfs")[0];
        this._cellXfsNode.removeAttribute("count");
    }

    createStyle(sourceId) {
        debug("createStyle: sourceId: %s", sourceId);

        let fontNode, borderNode, xfNode;
        if (sourceId) {
            debug("Cloning source xf node");
            const sourceXfNode = this._cellXfsNode.childNodes[sourceId];
            xfNode = sourceXfNode.cloneNode(true);

            if (sourceXfNode.getAttribute("applyFont") === "1") {
                const fontId = parseInt(sourceXfNode.getAttribute("fontId"));
                debug("Cloning source font node: %s", fontId);
                fontNode = this._fontsNode.childNodes[fontId].cloneNode(true);
            }

            if (sourceXfNode.getAttribute("applyBorder") === "1") {
                const borderId = parseInt(sourceXfNode.getAttribute("borderId"));
                debug("Cloning source font node: %s", borderId);
                borderNode = this._bordersNode.childNodes[borderId].cloneNode(true);
            }
        }

        if (!fontNode) fontNode = this._xml.createElement("font");
        this._fontsNode.appendChild(fontNode);

        if (!borderNode) borderNode = this._xml.createElement("border");
        this._bordersNode.appendChild(borderNode);

        if (!xfNode) xfNode = this._xml.createElement("xf");
        xfNode.setAttribute("fontId", this._fontsNode.childNodes.length - 1);
        xfNode.setAttribute("borderId", this._bordersNode.childNodes.length - 1);
        xfNode.setAttribute("applyFont", 1);
        xfNode.setAttribute("applyBorder", 1);
        this._cellXfsNode.appendChild(xfNode);

        return new Style(this, this._cellXfsNode.childNodes.length - 1, xfNode, fontNode, borderNode);
    }

    getNumberFormatCode(id) {
        debug("getNumberFormatCode(%o)", arguments);
        return this._numberFormatCodesById[id];
    }

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

    toString() {
        if (!this._numFmtsNode.childNodes.length) {
            this._xml.documentElement.removeChild(this._numFmtsNode);
        }

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
