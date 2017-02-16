"use strict";

// TODO: Tests
// TODO: JSDoc
// TODO: Debugs
// TODO: When populating a *new* cell you need to copy the style ref from the column or row (not sure what happens if both set)

const Range = require("./Range");
const Group = require("./Group");
const ArgHandler = require("./ArgHandler");
const addressConverter = require("./addressConverter");
const utils = require("./utils");
const debug = require("./debug")("Cell");
const xmlq = require("./xmlq");

/**
 * A cell
 */
class Cell {
    constructor(row, node) {
        this._row = row;
        this._initNode(node);
    }

    _initNode(node) {
        this._node = node;

        this._ref = addressConverter.fromAddress(this._node.attributes.r);
        
        const sharedFormulaId = this._getSharedFormulaId();
        this.sheet().updateMaxSharedFormulaId(sharedFormulaId);

        // This is a blunt way to make sure formula values get updated.
        // It just clears any stored values in case the referenced cell values change.
        if (xmlq.hasChild(this._node, 'f')) {
            xmlq.removeChild(this._node, 'v');
        }
    }

    tap(callback) {
        callback(this);
        return this;
    }

    thru(callback) {
        return callback(this);
    }

    // TODO: Test Range
    rangeTo(cell) {
        return new Range(this, cell);
    }

    // TODO: Test Group
    groupWith() {
        return new Group(this, arguments);
    }

    /**
     * Get the address of the column.
     * @param {{}} [opts] - Options
     * @param {boolean} [opts.includeSheetName] - Include the sheet name in the address.
     * @param {boolean} [opts.anchored] - Anchor the address.
     * @returns {string} The address
     */
    address(opts) {
        debug("address(%o)", arguments);
        const ref = this._ref;
        return addressConverter.toAddress({
            type: 'cell',
            rowNumber: ref.rowNumber,
            columnName: ref.columnName,
            sheetName: opts && opts.includeSheetName && this.sheet().name(),
            rowAnchored: opts && opts.rowAnchored,
            columnAnchored: opts && opts.columnAnchored
        });
    }

    /**
     * Clears the contents from the cell.
     * @returns {Cell} The cell.
     */
    clear() {
        debug("clear(%o)", arguments);

        // TODO: Move shared formula to some other cell. This would require parsing the formula... Push to v1.1?
        const sharedFormulaId = this._getSharedFormulaId();

        this._node.children = [];
        delete this._node.attributes.t;

        if (sharedFormulaId >= 0) this.sheet().clearCellsUsingSharedFormula(sharedFormulaId);

        return this;
    }

    /**
     * Gets the column name of the cell.
     * @returns {number} The column name.
     */
    columnName() {
        return this._ref.columnName;
    }

    /**
     * Gets the column number of the cell (1-based).
     * @returns {number} The column number.
     */
    columnNumber() {
        return this._ref.columnNumber;
    }

    /**
     * Returns a cell with a relative position given the offsets provided.
     * @param {number} rowOffset - The row offset (0 for the current row).
     * @param {number} columnOffset - The column offset (0 for the current column).
     * @returns {Cell} The relative cell.
     */
    relativeCell(rowOffset, columnOffset) {
        const row = rowOffset + this.rowNumber();
        const column = columnOffset + this.columnNumber();
        return this.sheet().cell(row, column);
    }

    /**
     * Gets the parent row of the cell.
     * @returns {Row} The parent row.
     */
    row() {
        return this._row;
    }

    column() {
        return this.sheet().column(this.columnNumber());
    }

    /**
     * Gets the row number of the cell (1-based).
     * @returns {number} The row number.
     */
    rowNumber() {
        return this._ref.rowNumber;
    }

    /**
     * Gets the parent sheet.
     * @returns {Sheet} The parent sheet.
     */
    sheet() {
        return this.row().sheet();
    }

    style() {
        debug("style(%o)", arguments);
        if (!this._style) {
            const styleId = this._node.attributes.s;
            this._style = this.workbook().styleSheet().createStyle(styleId); // TODO: Use proper method
            this._node.attributes.s = this._style.id();
        }

        return new ArgHandler("Cell.style")
            .case('string', name => {
                // Get single value
                return this._style.style(name);
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
                // Set a single value for all cells to a single value
                this._style.style(name, value);
                return this;
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
            .parse(arguments);
    }

    find(pattern, replacement) {
        pattern = utils.getRegExpForSearch(pattern);

        const value = this.value();
        if (typeof value !== 'string') return false;

        if (arguments.length === 2) {
            const replaced = value.replace(pattern, replacement);
            if (replaced === value) return false;
            this.value(replaced);
            return true;
        } else {
            return pattern.test(value);
        }
    }

    /**
     * Gets or sets the value of the cell.
     * @param {string|boolean|number|Date|null|undefined} [value] - The value to set.
     * @returns {string|boolean|number|Date|null|Cell} The value of the cell or the cell if setting.
     */
    value(value) {
        debug("value(%o)", arguments);
        return new ArgHandler('Cell.value')
            .case(() => {
                // Getter
                const type = this._node.attributes.t;

                // TODO: Validate node names.
                if (type === "s") {
                    const sharedIndex = this._node.children[0].children[0];
                    value = this.workbook().sharedStrings().getStringByIndex(sharedIndex);
                } else if (type === "inlineStr") {
                    value = this._node.children[0].children[0].children[0];
                } else if (type === "b") {
                    value = this._node.children[0].children[0] === 1;
                } else {
                    value = this._node.children[0] && this._node.children[0].children[0];
                }

                return value;
            })
            .case('*', value => {
                // Setter
                this.clear();

                let type, text;
                if (typeof value === "string") {
                    type = "s";
                    text = this.workbook().sharedStrings().getIndexForString(value);
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
            })
            .parse(arguments);
    }

    formula() {
        debug("formula(%o)", arguments);
        return new ArgHandler('Cell.formula')
            .case(() => {
                const fNode = xmlq.findChild(this._node, 'f');
                if (!fNode) return;

                // TODO: Return translated formula. Perhaps in v1.1?
                if (fNode.attributes.t === "shared" && !fNode.attributes.ref) return "SHARED";

                return fNode.children[0];
            })
            .case('string', formula => {
                this.clear();
                const fNode = { name: 'f', attributes: {}, children: [formula] };
                xmlq.appendChild(this._node, fNode);
                return this;
            });
    }

    // @ignore
    sharesFormula(sharedFormulaId) {
        debug("_sharesFormula(%o)", arguments);
        const fNode = xmlq.findChild(this._node, 'f');
        return fNode && fNode.attributes.si === sharedFormulaId;
    }

    _getSharedFormulaId() {
        debug("_getSharedFormulaRefId(%o)", arguments);
        const fNode = xmlq.findChild(this._node, 'f');
        return fNode && fNode.attributes.ref ? fNode.attributes.si : -1;
    }

    // @ignore
    setSharedFormula(formula, sharedIndex, sharedRef) {
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
        return this.row().workbook();
    }

    // @ignore
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
