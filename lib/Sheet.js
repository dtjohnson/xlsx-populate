"use strict";

const DOMParser = require('xmldom').DOMParser;
const parser = new DOMParser();
const Row = require("./Row");
const Column = require("./Column");
const Range = require("./Range");
const _ = require("lodash");
const debug = require("./debug")("Sheet");

const utils = require("./utils");
const xpath = require("./xpath");

class Sheet {
    constructor(workbook, idNode, text) {
        debug("constructor(...)");

        this._maxSharedFormulaId = -1;
        this._workbook = workbook;
        this._idNode = idNode;
        this._xml = parser.parseFromString(text);

        // This is a blunt way to make sure formula values get updated.
        // It just clears all stored values in case the referenced cell values change.
        const valueNodes = xpath("sml:sheetData/sml:row/sml:c/sml:f/../sml:*[name(.) !='f']", this._xml);
        valueNodes.forEach(valueNode => valueNode.parentNode.removeChild(valueNode));

        this._rows = [];
        this._sheetDataNode = this._xml.documentElement.getElementsByTagName("sheetData")[0];
        const rowNodes = this._sheetDataNode.childNodes;
        for (let i = 0; i < rowNodes.length; i++) {
            const rowNode = rowNodes[i];
            const row = new Row(this, rowNode);
            this._rows[row.rowNumber()] = row;
        }

        this._columns = [];
        this._colsNode = this._xml.documentElement.getElementsByTagName("cols")[0];
        if (!this._colsNode) {
            this._colsNode = this._xml.createElement("cols");

            // Must come before sheetData
            this._xml.documentElement.insertBefore(this._colsNode, this._sheetDataNode);
        }

        const colNodes = this._colsNode.childNodes;
        for (let i = colNodes.length - 1; i >= 0; i--) {
            const colNode = colNodes[i];
            const min = parseInt(colNode.getAttribute("min"));
            const max = parseInt(colNode.getAttribute("max"));

            for (let columnNumber = min; columnNumber <= max; columnNumber++) {
                const clonedColNode = colNode.cloneNode(true);
                clonedColNode.setAttribute("min", columnNumber);
                clonedColNode.setAttribute("max", columnNumber);
                this._columns[columnNumber] = new Column(this, clonedColNode);
            }

            this._colsNode.removeChild(colNode);
        }
    }

    workbook() {
        return this._workbook;
    }

    /**
     * Gets the cell with the given address.
     * @param {string} address - The address of the cell.
     * @returns {Cell} The cell.
     *//**
     * Gets the cell with the given row and column numbers.
     * @param {number} rowNumber - The row number of the cell.
     * @param {number} columnNumber - The column number of the cell.
     * @returns {Cell} The cell.
     */
    cell() {
        let rowNumber, columnNumber;
        if (arguments.length === 1) {
            const address = arguments[0];
            const ref = utils.addressToRowAndColumn(address);
            rowNumber = ref.row;
            columnNumber = ref.column;
        } else {
            rowNumber = arguments[0];
            columnNumber = arguments[1];
        }

        return this.row(rowNumber).cell(columnNumber);
    }

    range() {
        if (arguments.length === 1) {
            const refs = arguments[0].split(":");
            return this.range(refs[0], refs[1]);
        } else if (arguments.length === 2) {
            let startCell = arguments[0];
            let endCell = arguments[1];

            if (typeof startCell === "string") {
                startCell = this.cell(startCell);
                return this.range(startCell, endCell);
            }

            if (typeof endCell === "string") {
                endCell = this.cell(endCell);
                return this.range(startCell, endCell);
            }

            // TODO: Error check

            return new Range(startCell, endCell);
        } else if (arguments.length === 4) {
            return this.range(this.cell(arguments[0], arguments[1]), this.cell(arguments[1], arguments[2]));
        } else {
            // Error
        }
    }

    activate() {

    }

    remove() {
        
    }

    name() {
        return this._idNode.getAttribute("name");
    }

    find(pattern) {
        pattern = utils.getRegExpForSearch(pattern);

        let matches = [];
        this._rows.forEach(row => {
            if (!row) return;
            matches = matches.concat(row.find(pattern));
        });

        return matches;
    }

    replace(pattern, replacement) {
        pattern = utils.getRegExpForSearch(pattern);

        let count = 0;
        this._rows.forEach(row => {
            if (!row) return;
            count += row.replace(pattern, replacement);
        });

        return count;
    }

    /**
     * Gets the row with the given number.
     * @param {number} rowNumber - The row number.
     * @returns {Row} The row with the given number.
     */
    row(rowNumber) {
        debug("row(%o)", arguments);
        if (this._rows[rowNumber]) return this._rows[rowNumber];
        const rowNode = this._xml.createElement("row");
        rowNode.setAttribute("r", rowNumber);
        const row = new Row(this, rowNode);
        this._rows[rowNumber] = row;
        return row;
    }

    column(columnNumber) {
        debug("column(%o)", arguments);

        if (typeof columnNumber === "string") columnNumber = utils.columnNameToNumber(columnNumber);

        if (this._columns[columnNumber]) return this._columns[columnNumber];
        const colNode = this._xml.createElement("col");
        colNode.setAttribute("min", columnNumber);
        colNode.setAttribute("max", columnNumber);
        const column = new Column(this, colNode);
        this._columns[columnNumber] = column;
        return column;
    }

    toString() {
        // TODO: Be smarter about this and only move those that need it.

        this._rows.forEach(row => {
            if (row) {
                this._sheetDataNode.appendChild(row._node);

                row._cells.forEach(cell => {
                    if (cell) row._node.appendChild(cell._node);
                });
            }
        });

        let hasCol = false;
        this._columns.forEach(column => {
            if (column) {
                hasCol = true;
                this._colsNode.appendChild(column._node);
            }
        });

        if (!hasCol) this._xml.documentElement.removeChild(this._colsNode);

        return this._xml.toString();
    }

    _clearCellsUsingSharedFormula(sharedFormulaId) {
        debug("_clearCellsUsingSharedFormula(%o)", arguments);
        this._rows.forEach(row => {
            if (!row) return;
            row._cells.forEach(cell => {
                if (!cell) return;
                if (cell._sharesFormula(sharedFormulaId)) cell.clear();
            });
        });
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

/*
cell()
    .style() // Makes a copy the first time (true flag would not copy)
        .font()
            .bold(true)
            .italic(true)
            .underline(true)
            .strikethrough(true)
            .superscript(true) // verticalAlign = "superscript" ("baseline")
            .subscript(true) // verticalAlign = "subscript"
            .size(12)
            .family("Arial")
            .color("FF0000")
            .style()
        .alignment()
            .horizontal("center") // left, right, center
            .vertical("center") // top, center, bottom
            .wrapText(true)
            .indent(1)
            .orientation(45)
            .angleCounterClockwise(true) // orientation = 45
            .angleClockwise(true) // orientation = 135
            .verticalText(true) // orientation = 255
            .rotateTextUp(true) // orientation = 90
            .rotateTextDown(true) // orientation = 180
            .style()
        .fill()
            .patternType("solid")
            .foregroundColor("FF0000")
            .backgoundColor("FF0000")
            .style()
        .border()
            .left()
                .borderStyle("thin")
                .border()
            .right()
                .borderStyle("medium")
                .border()
            .bottom()
                .borderStyle("double")
        .numberFormat(43)

cell()
    .style() // Makes a copy the first time (true flag would not copy)
        .bold(true)
        .italic(true)
        .underline(true)
        .strikethrough(true)
        .superscript(true) // verticalAlign = "superscript" ("baseline")
        .subscript(true) // verticalAlign = "subscript"
        .fontSize(12)
        .fontFamily("Arial")
        .fontColor("FF0000")
        .horizontalAlignment("center") // left, right, center
        .verticalAlignment("center") // top, center, bottom
        .wrapText(true)
        .indent(1)
        .textOrientation(45)
        .angleTextCounterClockwise(true) // orientation = 45
        .angleTextClockwise(true) // orientation = 135
        .verticalText(true) // orientation = 255
        .rotateTextUp(true) // orientation = 90
        .rotateTextDown(true) // orientation = 180
        .fillPatternType("solid")
        .fillForegroundColor("FF0000")
        .fillBackgoundColor("FF0000")
        .leftBorderStyle("thin")
        .rightBorderStyle("medium")
        .bottomBorderStyle("double")
        .borderColor("FF0000")
        .numberFormat(43);

/*
font requires applyFont="1"

Bold:
style -> font with <b/>

Italic:
style -> font with <i/>

Underline:
style -> font with <u/>

Strikethrough:
style -> font with <strike/>

Superscript:
style -> font with <vertAlign val="superscript"/>

Subscript:
style -> font with <vertAlign val="subscript"/>

Font size x:
style -> font with <sz val="x"/>

Font family x:
style -> font with <name val="x"/>

Alignmen:
style with applyAlignment="1" and child <alignment horizontal="left|center|right" vertical="left|center|right" wrapText="1" indent="1+" textRotation="45"/>
Angle Counter Clockwise = 45
Angle Clockwise = 135
Vertical text = 255
Rotate Text Up = 90
Rotate Text Down = 180

Fill:
style with applyFill="1" -> fill with:
<patternFill patternType="solid">
    <fgColor rgb="FFFFFF00"/>
    <bgColor indexed="64"/>
</patternFill>

Font color x:
style -> font with <color rgb="x"/> indexed, theme, or rgb

Border:
style with applyBorder="1" -> border:
<border>
    <left /> style="thin|medium|thick|double"
    child <color> indexed = 64 is black (not sure if necessary)
    ...
</border>

Number Format:
style -> numFmtId
Create numFmt if not standard: http://polymathprogrammer.com/2011/02/15/built-in-styles-for-excel-open-xml/
 */