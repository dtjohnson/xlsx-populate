"use strict";

// TODO: Tests
// TODO: JSDoc

const _ = require("lodash");
const Row = require("./Row");
const Column = require("./Column");
const Range = require("./Range");
const debug = require("./debug")("Sheet");

const xmlq = require("./xmlq");
const utils = require("./utils");

class Sheet {
    constructor(workbook, idNode, node) {
        debug("constructor(...)");

        this._maxSharedFormulaId = -1;
        this._workbook = workbook;
        this._idNode = idNode;
        this._node = node;

        // Create the rows.
        this._rows = [];
        this._sheetDataNode = xmlq.findChild(this._node, "sheetData");
        this._sheetDataNode.children.forEach(rowNode => {
            const row = new Row(this, rowNode);
            this._rows[row.rowNumber()] = row;
        });

        // Create the columns.
        this._columns = [];
        this._colsNode = xmlq.findChild(this._node, "cols");
        if (!this._colsNode) {
            this._colsNode = { name: 'cols', children: [] };
        }

        // Excel will merge columns using min/max. Break them apart.
        const colNodes = this._colsNode.children;
        this._colsNode.children = [];
        colNodes.forEach(colNode => {
            const min = colNode.attributes.min;
            const max = colNode.attributes.max;

            for (let columnNumber = min; columnNumber <= max; columnNumber++) {
                const clonedColNode = _.cloneDeep(colNode);
                xmlq.appendChild(this._colsNode, clonedColNode);
                clonedColNode.attributes.min = columnNumber;
                clonedColNode.attributes.max = columnNumber;
                this._columns[columnNumber] = new Column(this, clonedColNode);
            }
        });

        this._mergeCells = {};
    }

    updateMaxSharedFormulaId(sharedFormulaId) {
        if (sharedFormulaId > this._maxSharedFormulaId) {
            this._maxSharedFormulaId = sharedFormulaId;
        }
    }

    // ignore
    mergeCells(address) {
        this._mergeCells[address] = {
            name: 'mergeCell',
            attributes: { ref: address }
        };
    }

    // ignore
    unergedCells(address) {
        delete this._mergeCells[address];
    }

    // ignore
    areCellsMerged(address) {
        return this._mergeCells.hasOwnProperty(address);
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
    // TODO: _ArgParser
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

    // TODO: _ArgParser
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
        // TODO
    }

    remove() {
        // TODO
    }

    // TODO: Set name
    name() {
        return this._idNode.attributes.name;
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
        const rowNode = { name: 'row', attributes: { r: rowNumber }, children: [] };
        const row = new Row(this, rowNode);
        this._rows[rowNumber] = row;
        return row;
    }

    column(columnNumber) {
        debug("column(%o)", arguments);

        if (typeof columnNumber === "string") columnNumber = utils.columnNameToNumber(columnNumber);

        if (this._columns[columnNumber]) return this._columns[columnNumber];
        const colNode = {
            name: 'col',
            attributes: {
                min: columnNumber,
                max: columnNumber
            }
        };
        const column = new Column(this, colNode);
        this._columns[columnNumber] = column;
        return column;
    }

    toObject() {
        // Rows must be in order.
        this._sheetDataNode.children = [];
        this._rows.forEach(row => {
            if (row) this._sheetDataNode.children.push(row.toObject());
        });

        // Columns must be in order.
        this._colsNode.children = [];
        this._columns.forEach(column => {
            if (column) this._colsNode.children.push(column.toObject());
        });

        // The cols node should only be present if there are columns defined.
        // If must also appear before the sheetData node.
        if (this._colsNode.children.length && !xmlq.hasChild(this._node, "cols")) {
            xmlq.insertBefore(this._node, this._colsNode, this._sheetDataNode);
        }

        if (_.isEmpty(this._mergeCells)) {
            xmlq.removeChild(this._node, "mergeCells");
        } else {
            xmlq.insertAfter(this._node, {
                name: 'mergeCells',
                children: _.values(this._mergeCells)
            }, this._sheetDataNode);
        }

        return this._node;
    }

    // @ignore
    clearCellsUsingSharedFormula(sharedFormulaId) {
        debug("_clearCellsUsingSharedFormula(%o)", arguments);
        this._rows.forEach(row => {
            if (!row) return;
            row._cells.forEach(cell => {
                if (!cell) return;
                if (cell.sharesFormula(sharedFormulaId)) cell.clear();
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
