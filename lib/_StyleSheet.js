"use strict";

// TODO: Tests
// TODO: JSDoc
// TODO: Debugs
// TODO: Enable number formats

const DOMParser = require('xmldom').DOMParser;
const parser = new DOMParser();
const debug = require("debug")("_StyleSheet");

const Style = require("./Style");

class _StyleSheet {
    constructor(text) {
        this._xml = parser.parseFromString(text);


        // this._numFmtsNode = this._xml.documentElement.getElementsByTagName("numFmts")[0];
        // if (this._numFmtsNode) {
        //     this._numFmtsNode.removeAttribute("count");
        // } else {
        //     this._numFmtsNode = this._xml.createElement("numFmts");
        //     this._xml.documentElement.appendChild(this._numFmtsNode);
        // }

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

        return new Style(this._cellXfsNode.childNodes.length - 1, xfNode, fontNode, borderNode);
    }

    toString() {
        return this._xml.toString();
    }
}

module.exports = _StyleSheet;


/*
xl/styles.xml

<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main">
	<fonts count="2" x14ac:knownFonts="1">
		<font>
			<sz val="11"/>
			<color theme="1"/>
			<name val="Calibri"/>
			<family val="2"/>
			<scheme val="minor"/>
		</font>
		<font>
			<b/>
			<sz val="11"/>
			<color theme="1"/>
			<name val="Calibri"/>
			<family val="2"/>
			<scheme val="minor"/>
		</font>
	</fonts>
	<fills count="2">
		<fill>
			<patternFill patternType="none"/>
		</fill>
		<fill>
			<patternFill patternType="gray125"/>
		</fill>
	</fills>
	<borders count="1">
		<border>
			<left/>
			<right/>
			<top/>
			<bottom/>
			<diagonal/>
		</border>
	</borders>
	<cellStyleXfs count="1">
		<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
	</cellStyleXfs>
	<cellXfs count="2">
		<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
		<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">
			<alignment horizontal="center"/>
		</xf>
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
