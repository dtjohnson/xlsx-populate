"use strict";

const ArgHandler = require("./ArgHandler");
const addressConverter = require("./addressConverter");
const dateConverter = require("./dateConverter");
const regexify = require("./regexify");
const debug = require("./debug")("Cell");
const xmlq = require("./xmlq");

/**
 * A cell
 */
class Cell {
    // /**
    //  * Creates a new instance of cell.
    //  * @param {Row} row - The parent row.
    //  * @param {{}} node - The cell node.
    //  */
    constructor(row, node) {
        debug('constructor(...)');
        this._row = row;
        this._init(node);
    }

    /**
     * Get the address of the column.
     * @param {{}} [opts] - Options
     * @param {boolean} [opts.includeSheetName] - Include the sheet name in the address.
     * @param {boolean} [opts.rowAnchored] - Anchor the row.
     * @param {boolean} [opts.columnAnchored] - Anchor the column.
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
     * Gets the parent column of the cell.
     * @returns {Column} The parent column.
     */
    column() {
        debug('column(%o)', arguments);
        return this.sheet().column(this.columnNumber());
    }

    /**
     * Clears the contents from the cell.
     * @returns {Cell} The cell.
     */
    clear() {
        debug("clear(%o)", arguments);

        // TODO in future version: Move shared formula to some other cell. This would require parsing the formula...
        const sharedFormulaId = this._getSharedFormulaRefId();

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
        debug('columnName(%o)', arguments);
        return this._ref.columnName;
    }

    /**
     * Gets the column number of the cell (1-based).
     * @returns {number} The column number.
     */
    columnNumber() {
        debug('columnNumber(%o)', arguments);
        return this._ref.columnNumber;
    }

    /**
     * Find the given pattern in the cell and optionally replace it.
     * @param {string|RegExp} pattern - The pattern to look for. Providing a string will result in a case-insensitive substring search. Use a RegExp for more sophisticated searches.
     * @param {string|function} [replacement] - The text to replace or a String.replace callback function. If pattern is a string, all occurrences of the pattern in the cell will be replaced.
     * @returns {boolean} A flag indicating if the pattern was found.
     */
    find(pattern, replacement) {
        debug('find(%o)', arguments);
        pattern = regexify(pattern);

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
     * Gets the formula in the cell. Note that if a formula was set as part of a range, the getter will return 'SHARED'. This is a limitation that may be addressed in a future release.
     * @returns {string} The formula in the cell.
     *//**
     * Sets the formula in the cell.
     * @param {string} formula - The formula to set.
     * @returns {Cell} The cell.
     */
    formula() {
        debug("formula(%o)", arguments);
        return new ArgHandler('Cell.formula')
            .case(() => {
                const fNode = xmlq.findChild(this._node, 'f');
                if (!fNode) return;

                // TODO in future: Return translated formula.
                if (fNode.attributes.t === "shared" && !fNode.attributes.ref) return "SHARED";

                return fNode.children[0];
            })
            .case('nil', () => {
                this.clear();
                return this;
            })
            .case('string', formula => {
                this.clear();
                const fNode = { name: 'f', attributes: {}, children: [formula] };
                xmlq.appendChild(this._node, fNode);
                return this;
            })
            .handle(arguments);
    }

    /**
     * Gets the hyperlink attached to the cell.
     * @returns {string|undefined} The hyperlink or undefined if not set.
     *//**
     * Set or clear the hyperlink on the cell.
     * @param {string|undefined} hyperlink - The hyperlink to set or undefined to clear.
     * @returns {Cell} The cell.
     */
    hyperlink() {
        debug("hyperlink(%o)", arguments);
        return new ArgHandler('Cell.hyperlink')
            .case(() => {
                return this.sheet().hyperlink(this.address());
            })
            .case('*', hyperlink => {
                this.sheet().hyperlink(this.address(), hyperlink);
                return this;
            })
            .handle(arguments);
    }

    /**
     * Callback used by tap.
     * @callback Cell~tapCallback
     * @param {Cell} cell - The cell
     * @returns {undefined}
     */
    /**
     * Invoke a callback on the cell and return the cell. Useful for method chaining.
     * @param {Cell~tapCallback} callback - The callback function.
     * @returns {Cell} The cell.
     */
    tap(callback) {
        debug('tap(%o)', arguments);
        callback(this);
        return this;
    }

    /**
     * Callback used by thru.
     * @callback Cell~thruCallback
     * @param {Cell} cell - The cell
     * @returns {*} The value to return from thru.
     */
    /**
     * Invoke a callback on the cell and return the value provided by the callback. Useful for method chaining.
     * @param {Cell~thruCallback} callback - The callback function.
     * @returns {*} The return value of the callback.
     */
    thru(callback) {
        debug('thru(%o)', arguments);
        return callback(this);
    }

    /**
     * Create a range from this cell and another.
     * @param {Cell|string} cell - The other cell or cell address to range to.
     * @returns {Range} The range.
     */
    rangeTo(cell) {
        debug('rangeTo(%o)', arguments);
        return this.sheet().range(this, cell);
    }

    /**
     * Returns a cell with a relative position given the offsets provided.
     * @param {number} rowOffset - The row offset (0 for the current row).
     * @param {number} columnOffset - The column offset (0 for the current column).
     * @returns {Cell} The relative cell.
     */
    relativeCell(rowOffset, columnOffset) {
        debug('relativeCell(%o)', arguments);
        const row = rowOffset + this.rowNumber();
        const column = columnOffset + this.columnNumber();
        return this.sheet().cell(row, column);
    }

    /**
     * Gets the parent row of the cell.
     * @returns {Row} The parent row.
     */
    row() {
        debug('row(%o)', arguments);
        return this._row;
    }

    /**
     * Gets the row number of the cell (1-based).
     * @returns {number} The row number.
     */
    rowNumber() {
        debug('rowNumber(%o)', arguments);
        return this._ref.rowNumber;
    }

    /**
     * Gets the parent sheet.
     * @returns {Sheet} The parent sheet.
     */
    sheet() {
        debug('sheet(%o)', arguments);
        return this.row().sheet();
    }

    /**
     * Gets an individual style.
     * @param {string} name - The name of the style.
     * @returns {*} The style.
     *//**
     * Gets multiple styles.
     * @param {Array.<string>} names - The names of the style.
     * @returns {object.<string, *>} Object whose keys are the style names and values are the styles.
     *//**
     * Sets an individual style.
     * @param {string} name - The name of the style.
     * @param {*} value - The value to set.
     * @returns {Cell} The cell.
     *//**
     * Sets multiple styles.
     * @param {object.<string, *>} styles - Object whose keys are the style names and values are the styles to set.
     * @returns {Cell} The cell.
     */
    style() {
        debug("style(%o)", arguments);
        if (!this._style) {
            const styleId = this._node.attributes.s;
            this._style = this.workbook().styleSheet().createStyle(styleId);
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
            .handle(arguments);
    }

    /**
     * Gets the value of the cell.
     * @returns {string|boolean|number|Date|undefined} The value of the cell.
     *//**
     * Sets the value of the cell.
     * @param {string|boolean|number|null|undefined} value - The value to set.
     * @returns {Cell} The cell.
     */
    value(value) {
        debug("value(%o)", arguments);
        return new ArgHandler('Cell.value')
            .case(() => {
                // Getter
                const type = this._node.attributes.t;

                if (type === "s") {
                    const sharedIndex = xmlq.findChild(this._node, 'v').children[0];
                    value = this.workbook().sharedStrings().getStringByIndex(sharedIndex);
                } else if (type === "inlineStr") {
                    value = xmlq.findChild(xmlq.findChild(this._node, 'is'), 't').children[0];
                } else if (type === "b") {
                    value = xmlq.findChild(this._node, 'v').children[0] === 1;
                } else {
                    const vNode = xmlq.findChild(this._node, 'v');
                    value = vNode && vNode.children[0];
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
                    text = dateConverter.dateToNumber(value);
                } else if (value) {
                    throw new Error("Cell.value: Unsupported value");
                } else {
                    return this;
                }

                if (type) this._node.attributes.t = type;
                const vNode = { name: 'v', children: [text] }; // Don't create attributes to save memory
                xmlq.appendChild(this._node, vNode);

                return this;
            })
            .handle(arguments);
    }

    /**
     * Gets the parent workbook.
     * @returns {Workbook} The parent workbook.
     */
    workbook() {
        debug('workbook(%o)', arguments);
        return this.row().workbook();
    }

    /**
     * Gets the formula if a shared formula ref cell.
     * @returns {string|undefined} The formula.
     * @ignore
     */
    getSharedRefFormula() {
        const fNode = xmlq.findChild(this._node, 'f');
        return fNode && fNode.attributes.ref && fNode.children[0];
    }

    /**
     * Check if this cell uses a given shared a formula ID.
     * @param {number} id - The shared formula ID.
     * @returns {boolean} A flag indicating if shared.
     * @ignore
     */
    sharesFormula(id) {
        debug("sharesFormula(%o)", arguments);
        const fNode = xmlq.findChild(this._node, 'f');
        return fNode && fNode.attributes.si === id;
    }

    /**
     * Set a shared formula on the cell.
     * @param {number} id - The shared formula index.
     * @param {string} [formula] - The formula (if the reference cell).
     * @param {string} [sharedRef] - The address of the shared range.
     * @returns {undefined}
     * @ignore
     */
    setSharedFormula(id, formula, sharedRef) {
        debug("setSharedFormula(%o)", arguments);
        this.clear();

        const fNode = {
            name: 'f',
            attributes: {
                t: 'shared',
                si: id
            },
            children: []
        };
        xmlq.appendChild(this._node, fNode);

        if (sharedRef) fNode.attributes.ref = sharedRef;
        if (formula) fNode.children = [formula];
    }

    /**
     * Convert the cell to an object.
     * @returns {{}} The object form.
     * @ignore
     */
    toObject() {
        debug('toObject(%o)', arguments);
        return this._node;
    }

    /**
     * Get the cell's shared formula ID if it is a shared formula reference cell.
     * @returns {number} The shared formula ID.
     * @private
     */
    _getSharedFormulaRefId() {
        debug("_getSharedFormulaRefId(%o)", arguments);
        const fNode = xmlq.findChild(this._node, 'f');
        return fNode && fNode.attributes.ref ? fNode.attributes.si : -1;
    }

    /**
     * Initialize the cell node.
     * @param {{}} [node] - The node
     * @returns {undefined}
     * @private
     */
    _init(node) {
        debug('_init(...)');
        this._node = node;

        this._ref = addressConverter.fromAddress(this._node.attributes.r);

        const sharedFormulaId = this._getSharedFormulaRefId();
        this.sheet().updateMaxSharedFormulaId(sharedFormulaId);

        // This is a blunt way to make sure formula values get updated.
        // It just clears any stored values in case the referenced cell values change.
        if (xmlq.hasChild(this._node, 'f')) {
            xmlq.removeChild(this._node, 'v');
        }
    }
}

module.exports = Cell;

/*
<c r="A6" s="1" t="s">
    <v>2</v>
</c>
*/
