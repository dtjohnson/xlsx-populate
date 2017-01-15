"use strict";

const DOMParser = require('xmldom').DOMParser;
const parser = new DOMParser();
const Row = require("./Row");

const xpath = require("./xpath");

class Sheet {
    constructor(workbook, idNode, text) {
        this._workbook = workbook;
        this._idNode = idNode;
        this._xml = parser.parseFromString(text);

        // This is a blunt way to make sure formula values get updated.
        // It just clears all stored values in case the referenced cell values change.
        const valueNodes = xpath("sml:sheetData/sml:row/sml:c/sml:f/../sml:*[name(.) !='f']", this._xml);
        valueNodes.forEach(valueNode => valueNode.parentNode.removeChild(valueNode));

        this._rows = [];
        const sheetDataNode = this._xml.documentElement.getElementsByTagName("sheetData")[0];
        const rowNodes = sheetDataNode.childNodes;
        for (let i = 0; i < rowNodes.length; i++) {
            const rowNode = rowNodes[i];
            const row = new Row(this, rowNode);
            this._rows[row.rowNumber()] = row;
        }
    }

    toString() {
        return this._xml.toString();
    }
}

module.exports = Sheet;

/*
xl/worksheets/sheetN.xml

<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
	<dimension ref="A1:I8"/>
	<sheetViews>
		<sheetView tabSelected="1" workbookViewId="0">
			<selection activeCell="A9" sqref="A9"/>
		</sheetView>
	</sheetViews>
	<sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
	<sheetData>
		<row r="1" spans="1:9" x14ac:dyDescent="0.25">
			<c r="A1" t="s">
				<v>0</v>
			</c>
		</row>
		<row r="2" spans="1:9" x14ac:dyDescent="0.25">
			<c r="A2" t="str">
				<f>A1</f><v>Foo</v>
			</c><c r="B2">
				<f t="shared" ref="B2:I2" si="0">B1</f><v>0</v>
			</c><c r="C2">
				<f t="shared" si="0"/>
				<v>0</v>
			</c><c r="D2">
				<f t="shared" si="0"/>
				<v>0</v>
			</c><c r="E2">
				<f t="shared" si="0"/>
				<v>0</v>
			</c><c r="F2">
				<f t="shared" si="0"/>
				<v>0</v>
			</c><c r="G2">
				<f t="shared" si="0"/>
				<v>0</v>
			</c><c r="H2">
				<f t="shared" si="0"/>
				<v>0</v>
			</c><c r="I2">
				<f t="shared" si="0"/>
				<v>0</v>
			</c>
		</row>
		<row r="3" spans="1:9" x14ac:dyDescent="0.25">
			<c r="A3" t="s">
				<v>1</v>
			</c><c r="B3" t="s">
				<v>1</v>
			</c><c r="C3" t="s">
				<v>1</v>
			</c><c r="D3" t="s">
				<v>1</v>
			</c><c r="E3" t="s">
				<v>1</v>
			</c><c r="F3" t="s">
				<v>1</v>
			</c><c r="G3" t="s">
				<v>1</v>
			</c><c r="H3" t="s">
				<v>1</v>
			</c><c r="I3" t="s">
				<v>1</v>
			</c>
		</row>
		<row r="4" spans="1:9" x14ac:dyDescent="0.25">
			<c r="A4">
				<v>1</v>
			</c><c r="B4">
				<v>2</v>
			</c><c r="C4">
				<v>3</v>
			</c><c r="D4">
				<v>4</v>
			</c><c r="E4">
				<v>5</v>
			</c><c r="F4">
				<v>6</v>
			</c><c r="G4">
				<v>7</v>
			</c><c r="H4">
				<v>8</v>
			</c><c r="I4">
				<v>9</v>
			</c>
		</row>
		<row r="6" spans="1:9" x14ac:dyDescent="0.25">
			<c r="A6" s="1" t="s">
				<v>2</v>
			</c><c r="B6" s="1"/>
			<c r="C6" s="1"/>
		</row>
		<row r="7" spans="1:9" x14ac:dyDescent="0.25">
			<c r="A7" t="s">
				<v>0</v>
			</c>
		</row>
		<row r="8" spans="1:9" x14ac:dyDescent="0.25">
			<c r="A8" t="s">
				<v>3</v>
			</c>
		</row>
	</sheetData>
	<mergeCells count="1">
		<mergeCell ref="A6:C6"/>
	</mergeCells>
	<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
	<pageSetup orientation="portrait" verticalDpi="300" r:id="rId1"/>
</worksheet>
*/
