"use strict";

const _ = require("lodash");
const Row = require("./Row");
const Column = require("./Column");
const Range = require("./Range");
const debug = require("./debug")("Sheet");
const xmlq = require("./xmlq");
const regexify = require("./regexify");
const addressConverter = require("./addressConverter");
const ArgHandler = require("./ArgHandler");

/**
 * A worksheet.
 */
class Sheet {
    // /**
    //  * Creates a new instance of Sheet.
    //  * @param {Workbook} workbook - The parent workbook.
    //  * @param {{}} idNode - The sheet ID node (from the parent workbook).
    //  * @param {{}} node - The sheet node.
    //  */
    constructor(workbook, idNode, node) {
        debug("constructor(...)");
        this._init(workbook, idNode, node);
    }

    /**
     * Gets the cell with the given address.
     * @param {string} address - The address of the cell.
     * @returns {Cell} The cell.
     *//**
     * Gets the cell with the given row and column numbers.
     * @param {number} rowNumber - The row number of the cell.
     * @param {string|number} columnNameOrNumber - The column name or number of the cell.
     * @returns {Cell} The cell.
     */
    cell() {
        debug("cell(%o)", arguments);
        return new ArgHandler('Sheet.cell')
            .case('string', address => {
                const ref = addressConverter.fromAddress(address);
                if (ref.type !== 'cell') throw new Error('Sheet.cell: Invalid address.');
                return this.row(ref.rowNumber).cell(ref.columnNumber);
            })
            .case(['number', '*'], (rowNumber, columnNameOrNumber) => {
                return this.row(rowNumber).cell(columnNameOrNumber);
            })
            .handle(arguments);
    }

    /**
     * Gets a column in the sheet.
     * @param {string|number} columnNameOrNumber - The name or number of the column.
     * @returns {Column} The column.
     */
    column(columnNameOrNumber) {
        debug("column(%o)", arguments);

        const columnNumber = typeof columnNameOrNumber === "string" ? addressConverter.columnNameToNumber(columnNameOrNumber) : columnNameOrNumber;

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

    /**
     * Gets a defined name scoped to the sheet.
     * @param {string} name - The defined name.
     * @returns {undefined|Cell|Range|Row|Column} The named selection or undefined if name not found.
     * @throws {Error} Will throw if address in defined name is not supported.
     */
    definedName(name) {
        return this.workbook().scopedDefinedName(name, this);
    }

    /**
     * Find the given pattern in the sheet and optionally replace it.
     * @param {string|RegExp} pattern - The pattern to look for. Providing a string will result in a case-insensitive substring search. Use a RegExp for more sophisticated searches.
     * @param {string|function} [replacement] - The text to replace or a String.replace callback function. If pattern is a string, all occurrences of the pattern in each cell will be replaced.
     * @returns {Array.<Cell>} The matching cells.
     */
    find(pattern, replacement) {
        debug("find(%o)", arguments);
        pattern = regexify(pattern);

        let matches = [];
        this._rows.forEach(row => {
            if (!row) return;
            matches = matches.concat(row.find(pattern, replacement));
        });

        return matches;
    }

    /**
     * Get the name of the sheet.
     * @returns {string} The sheet name.
     */
    name() {
        debug("name(%o)", arguments);
        return this._idNode.attributes.name;
    }

    /**
     * Gets a range from the given range address.
     * @param {string} address - The range address (e.g. 'A1:B3').
     * @returns {Range} The range.
     *//**
     * Gets a range from the given cells or cell addresses.
     * @param {string|Cell} startCell - The starting cell or cell address (e.g. 'A1').
     * @param {string|Cell} endCell - The ending cell or cell address (e.g. 'B3').
     * @returns {Range} The range.
     *//**
     * Gets a range from the given row numbers and column names or numbers.
     * @param {number} startRowNumber - The starting cell row number.
     * @param {string|number} startColumnNameOrNumber - The starting cell column name or number.
     * @param {number} endRowNumber - The ending cell row number.
     * @param {string|number} endColumnNameOrNumber - The ending cell column name or number.
     * @returns {Range} The range.
     */
    range() {
        debug("range(%o)", arguments);
        return new ArgHandler('Cell.range')
            .case('string', address => {
                const ref = addressConverter.fromAddress(address);
                if (ref.type !== 'range') throw new Error('Sheet.range: Invalid address');
                return this.range(ref.startRowNumber, ref.startColumnNumber, ref.endRowNumber, ref.endColumnNumber);
            })
            .case(['*', '*'], (startCell, endCell) => {
                if (typeof startCell === "string") startCell = this.cell(startCell);
                if (typeof endCell === "string") endCell = this.cell(endCell);
                return new Range(startCell, endCell);
            })
            .case(['number', '*', 'number', '*'], (startRowNumber, startColumnNameOrNumber, endRowNumber, endColumnNameOrNumber) => {
                return this.range(this.cell(startRowNumber, startColumnNameOrNumber), this.cell(endRowNumber, endColumnNameOrNumber));
            })
            .handle(arguments);
    }

    /**
     * Gets the row with the given number.
     * @param {number} rowNumber - The row number.
     * @returns {Row} The row with the given number.
     */
    row(rowNumber) {
        debug("row(%o)", arguments);
        if (this._rows[rowNumber]) return this._rows[rowNumber];

        const rowNode = {
            name: 'row',
            attributes: {
                r: rowNumber
            },
            children: []
        };

        const row = new Row(this, rowNode);
        this._rows[rowNumber] = row;
        return row;
    }

    /**
     * Get the range of cells in the sheet that have contained a value or style at any point. Useful for extracting the entire sheet contents.
     * @returns {Range|undefined} The used range or undefined if no cells in the sheet are used.
     */
    usedRange() {
        debug("usedRange(%o)", arguments);
        const minRowNumber = _.findIndex(this._rows);
        const maxRowNumber = this._rows.length - 1;

        let minColumnNumber = 0;
        let maxColumnNumber = 0;
        for (let i = 0; i < this._rows.length; i++) {
            const row = this._rows[i];
            if (!row) continue;

            const minUsedColumnNumber = row.minUsedColumnNumber();
            const maxUsedColumnNumber = row.maxUsedColumnNumber();
            if (minUsedColumnNumber > 0 && (!minColumnNumber || minUsedColumnNumber < minColumnNumber)) minColumnNumber = minUsedColumnNumber;
            if (maxUsedColumnNumber > 0 && (!maxColumnNumber || maxUsedColumnNumber > maxColumnNumber)) maxColumnNumber = maxUsedColumnNumber;
        }

        // Return undefined if nothing in the sheet is used.
        if (minRowNumber <= 0 || minColumnNumber <= 0 || maxRowNumber <= 0 || maxColumnNumber <= 0) return;

        return this.range(minRowNumber, minColumnNumber, maxRowNumber, maxColumnNumber);
    }

    /**
     * Gets the parent workbook.
     * @returns {XlsxPopulate} The parent workbook.
     */
    workbook() {
        debug("workbook(%o)", arguments);
        return this._workbook;
    }

    /**
     * Get a value indicating whether the cells in the given address are merged.
     * @param {string} address - The address to check.
     * @returns {boolean} True if merged, false if not merged.
     * @ignore
     */
    areCellsMerged(address) {
        debug("areCellsMerged(%o)", arguments);
        return this._mergeCells.hasOwnProperty(address);
    }

    /**
     * Clear cells that are using a given shared formula ID.
     * @param {number} sharedFormulaId - The shared formula ID.
     * @returns {undefined}
     * @ignore
     */
    clearCellsUsingSharedFormula(sharedFormulaId) {
        debug("clearCellsUsingSharedFormula(%o)", arguments);
        this._rows.forEach(row => {
            if (!row) return;
            row.clearCellsUsingSharedFormula(sharedFormulaId);
        });
    }

    /**
     * Get an existing column style ID.
     * @param {number} columnNumber - The column number.
     * @returns {undefined|number} The style ID.
     * @ignore
     */
    existingColumnStyleId(columnNumber) {
        debug("existingColumnStyleId(%o)", arguments);
        return this._columns[columnNumber] && this._columns[columnNumber].styleId();
    }

    /**
     * Increment and return the max shared formula ID.
     * @returns {number} The new max shared formula ID.
     * @ignore
     */
    incrementMaxSharedFormulaId() {
        debug("incrementMaxSharedFormulaId(%o)", arguments);
        return ++this._maxSharedFormulaId;
    }

    /**
     * Merge cells by adding a mergeCell entry.
     * @param {string} address - The address to merge.
     * @returns {undefined}
     * @ignore
     */
    mergeCells(address) {
        debug("mergeCells(%o)", arguments);
        this._mergeCells[address] = {
            name: 'mergeCell',
            attributes: { ref: address },
            children: []
        };
    }

    /**
     * Convert the sheet to an object.
     * @returns {{}} The object form.
     * @ignore
     */
    toObject() {
        debug("toObject(%o)", arguments);

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

        // The mergeCells node must be after the sheetData node and should not be present unless it has children.
        this._mergeCellsNode.children = _.values(this._mergeCells);
        if (this._mergeCellsNode.children.length) {
            xmlq.insertAfter(this._node, this._mergeCellsNode, this._sheetDataNode);
        } else {
            xmlq.removeChild(this._node, this._mergeCellsNode);
        }

        return this._node;
    }

    /**
     * Update the max shared formula ID to the given value if greater than current.
     * @param {number} sharedFormulaId - The new shared formula ID.
     * @returns {undefined}
     * @ignore
     */
    updateMaxSharedFormulaId(sharedFormulaId) {
        debug("updateMaxSharedFormulaId(%o)", arguments);
        if (sharedFormulaId > this._maxSharedFormulaId) {
            this._maxSharedFormulaId = sharedFormulaId;
        }
    }

    /**
     * Unmerges cells by removing a mergeCell entry.
     * @param {string} address - The address to unmerge.
     * @returns {undefined}
     * @ignore
     */
    unmergeCells(address) {
        debug("unmergeCells(%o)", arguments);
        delete this._mergeCells[address];
    }

    /**
     * Initializes the sheet.
     * @param {XlsxPopulate} workbook - The parent workbook.
     * @param {{}} idNode - The sheet ID node (from the parent workbook).
     * @param {{}} node - The sheet node.
     * @returns {undefined}
     * @private
     */
    _init(workbook, idNode, node) {
        debug("_init(%o)", arguments);
        this._workbook = workbook;
        this._idNode = idNode;
        this._node = node;
        this._maxSharedFormulaId = -1;
        this._mergeCells = {};

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
            this._colsNode = { name: 'cols', attributes: {}, children: [] };
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

        // Create the merge cells.
        this._mergeCellsNode = xmlq.findChild(this._node, "mergeCells");
        if (!this._mergeCellsNode) {
            this._mergeCellsNode = { name: 'mergeCells', attributes: {}, children: [] };
        }

        const mergeCellNodes = this._mergeCellsNode.children;
        this._mergeCellsNode.children = [];
        mergeCellNodes.forEach(mergeCellNode => {
            this._mergeCells[mergeCellNode.attributes.ref] = mergeCellNode;
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
