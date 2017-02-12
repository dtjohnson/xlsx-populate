"use strict";

// TODO: Tests
// TODO: JSDoc
// TODO: Debugs
// TODO: When populating a *new* cell you need to copy the style ref from the column or row (not sure what happens if both set)

const _ArgParser = require("./_ArgParser");
const utils = require("./utils");
const debug = require("./debug")("Cell");
const xmlq = require("./xmlq");

/**
 * A cell
 */
class Cell {
    constructor(row, node) {
        this._row = row;
        this._node = node;

        const sharedFormulaId = this._getSharedFormulaRefId();
        if (sharedFormulaId >= 0 && sharedFormulaId > this.sheet()._maxSharedFormulaId) {
            this.sheet()._mexSharedFormulaId = sharedFormulaId;
        }

        // This is a blunt way to make sure formula values get updated.
        // It just clears any stored values in case the referenced cell values change.
        if (xmlq.hasChild(this._node, 'f')) {
            xmlq.removeChild(this._node, 'v');
        }
    }

    call(handler) {
        handler(this);
        return this;
    }

    activate() {
        // TODO
    }

    rangeTo() {
        // TODO
    }

    groupWith() {
        // TODO
    }

    /**
     * Gets the address of the cell (e.g. "A5").
     * @returns {string} The cell address.
     */
    address() {
        if (arguments.length > 0) throw new Error('Cell.address: Cannot be set.');
        return this._node.attributes.r;
    }

    /**
     * Clears the contents from the cell.
     * @returns {Cell} The cell.
     */
    clear() {
        debug("clear(%o)", arguments);

        // TODO: Move shared formula to some other cell. This would require parsing the formula... Push to v1.1?
        const sharedFormulaId = this._getSharedFormulaRefId();

        this._node.children = [];
        delete this._node.attributes.t;

        if (sharedFormulaId >= 0) this.sheet()._clearCellsUsingSharedFormula(sharedFormulaId);

        return this;
    }

    /**
     * Gets the column name of the cell.
     * @returns {number} The column name.
     */
    columnName() {
        if (arguments.length > 0) throw new Error('Cell.columnName: Cannot be set.');
        return utils.columnNumberToName(this.columnNumber());
    }

    /**
     * Gets the column number of the cell (1-based).
     * @returns {number} The column number.
     */
    columnNumber() {
        if (arguments.length > 0) throw new Error('Cell.columnNumber: Cannot be set.');
        return utils.addressToRowAndColumn(this.address()).column;
    }

    /**
     * Gets the full address of the cell including sheet (e.g. "Sheet1!A5").
     * @returns {string} The full address.
     */
    fullAddress() {
        if (arguments.length > 0) throw new Error('Cell.fullAddress: Cannot be set.');
        return utils.addressToFullAddress(this.sheet().name(), this.address());
    }

    /**
     * Returns a cell with a relative position given the offsets provided.
     * @param {number} rowOffset - The row offset (0 for the current row).
     * @param {number} columnOffset - The column offset (0 for the current column).
     * @returns {Cell} The relative cell.
     */
    // TODO: Rename to offset like interop? https://msdn.microsoft.com/en-us/library/microsoft.office.interop.excel.range.offset.aspx
    relativeCell(rowOffset, columnOffset) {
        if (arguments.length !== 2) throw new Error("Cell.relativeCell: Invalid number of arguments");
        if (!Number.isInteger(rowOffset)) throw new Error("Cell.relativeCell: Invalid row offset");
        if (!Number.isInteger(columnOffset)) throw new Error("Cell.relativeCell: Invalid column offset");

        const row = rowOffset + this.rowNumber();
        const column = columnOffset + this.columnNumber();

        if (row < 1) throw new Error("Cell.relativeCell: Relative cell row position is less than 1");
        if (column < 1) throw new Error("Cell.relativeCell: Relative cell column position is less than 1");

        return this.sheet().cell(row, column);
    }

    /**
     * Gets the parent row of the cell.
     * @returns {Row} The parent row.
     */
    row() {
        if (arguments.length > 0) throw new Error('Cell.row: Cannot be set.');
        return this._row;
    }

    column() {
        // TODO
    }

    /**
     * Gets the row number of the cell (1-based).
     * @returns {number} The row number.
     */
    rowNumber() {
        if (arguments.length > 0) throw new Error('Cell.rowNumber: Cannot be set.');
        return utils.addressToRowAndColumn(this.address()).row;
    }

    /**
     * Gets the parent sheet.
     * @returns {Sheet} The parent sheet.
     */
    sheet() {
        if (arguments.length > 0) throw new Error('Cell.sheet: Cannot be set.');
        return this.row().sheet();
    }

    style() {
        debug("style(%o)", arguments);
        if (!this._style) {
            const styleId = this._node.attributes.s;
            this._style = this.workbook()._styleSheet.createStyle(styleId);
            this._node.attributes.s = this._style._id;
        }

        return new _ArgParser("Cell.style")
            .case(String, name => {
                // Get single value
                return this._style.style(name);
            })
            .case(Array, names => {
                // Get list of values
                const values = {};
                names.forEach(name => {
                    values[name] = this.style(name);
                });

                return values;
            })
            .case([String, undefined], (name, value) => {
                // Set a single value for all cells to a single value
                this._style.style(name, value);
                return this;
            })
            .case(Object, nameValues => {
                // Object of key value pairs to set
                for (const name in nameValues) {
                    if (!nameValues.hasOwnProperty(name)) continue;
                    const value = nameValues[name];
                    this.style(name, value);
                }

                return this;
            })
            .parse(arguments);
    }

    find(pattern) {
        pattern = utils.getRegExpForSearch(pattern);

        const value = this.value();
        if (typeof value !== 'string') return false;

        return pattern.test(value);
    }

    replace(pattern, replacement) {
        pattern = utils.getRegExpForSearch(pattern);

        const value = this.value();
        if (typeof value !== 'string') return false;

        const replaced = value.replace(pattern, replacement);
        if (replaced === value) return false;

        this.value(replaced);
        return true;
    }

    /**
     * Gets or sets the value of the cell.
     * @param {string|boolean|number|Date|null|undefined} [value] - The value to set.
     * @returns {string|boolean|number|Date|null|Cell} The value of the cell or the cell if setting.
     */
    // TODO: Switch to xq and _ArgParser
    value(value) {
        debug("value(%o)", arguments);

        if (arguments.length === 0) {
            // Getter
            const type = this._node.attributes.t;

            // TODO: Validate node names.
            if (type === "s") {
                const sharedIndex = this._node.children[0].children[0];
                value = this.workbook()._sharedStrings.getStringByIndex(sharedIndex);
            } else if (type === "inlineStr") {
                value = this._node.children[0].children[0].children[0];
            } else if (type === "b") {
                value = this._node.children[0].children[0] === 1;
            } else {
                value = this._node.children[0].children[0];
            }

            // TODO: Date

            return value;
        } else if (arguments.length === 1) {
            // Setter
            this.clear();

            let type, text;
            if (typeof value === "string") {
                type = "s";
                text = this.workbook()._sharedStrings.getIndexForString(value);
            } else if (typeof value === "boolean") {
                type = "b";
                text = value ? 1 : 0;
            } else if (typeof value === "number") {
                text = value;
            } else if (value instanceof Date) {
                text = utils.dateToExcelNumber(value);
            } else if (value) {
                throw new Error("Cell.value: Unsupported value");
            } else {
                return this;
            }

            if (type) this._node.attributes.t = type;
            const vNode = { name: 'v', attributes: {}, children: [text] };
            xmlq.appendChild(this._node, vNode);

            return this;
        } else {
            throw new Error("Cell.value: Unexpected number of arguments");
        }
    }

    // TODO: Switch to xq and _ArgParse
    formula(formula) {
        debug("formula(%o)", arguments);

        if (arguments.length === 0) {
            const fNode = xmlq.findChild(this._node, 'f');
            if (!fNode) return;

            // TODO: Return translated formula. Perhaps in v1.1?
            if (fNode.attributes.t === "shared" && !fNode.attributes.ref) return "SHARED";

            return fNode.children[0];
        } else if (arguments.length === 1) {
            this.clear();
            const fNode = { name: 'f', attributes: {}, children: [formula] };
            xmlq.appendChild(this._node, fNode);
            return this;
        } else {
            throw new Error();
        }
    }

    // TODO: xq
    _sharesFormula(sharedFormulaId) {
        debug("_sharesFormula(%o)", arguments);
        const fNode = xmlq.findChild(this._node, 'f');
        return fNode && fNode.attributes.si === sharedFormulaId;
    }

    // TODO: xq
    _getSharedFormulaRefId() {
        debug("_getSharedFormulaRefId(%o)", arguments);
        const fNode = xmlq.findChild(this._node, 'f');
        return fNode && fNode.attributes.ref ? fNode.attributes.si : -1;
    }

    // TODO: xq
    _setSharedFormula(formula, sharedIndex, sharedRef) {
        debug("_setFormula(%o)", arguments);
        this.clear();

        const fNode = {
            name: 'f',
            attributes: {
                t: 'shared',
                si: sharedIndex
            },
            children: []
        };
        xmlq.appendChild(this._node, fNode);

        if (sharedRef) fNode.attributes.ref = sharedRef;
        if (formula) fNode.children = [formula];
    }

    /**
     * Gets the parent workbook.
     * @returns {Workbook} The parent workbook.
     */
    workbook() {
        if (arguments.length > 0) throw new Error('Cell.workbook: Cannot be set.');
        return this.row().workbook();
    }

    toObject() {
        return this._node;
    }
}

module.exports = Cell;

/*
<c r="A6" s="1" t="s">
    <v>2</v>
</c>
*/
