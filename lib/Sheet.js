"use strict";

const _ = require("lodash");
const Cell = require("./Cell");
const Row = require("./Row");
const Column = require("./Column");
const Range = require("./Range");
const Relationships = require("./Relationships");
const xmlq = require("./xmlq");
const regexify = require("./regexify");
const addressConverter = require("./addressConverter");
const ArgHandler = require("./ArgHandler");
const colorIndexes = require("./colorIndexes");

// Order of the nodes as defined by the spec.
const nodeOrder = [
    "sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData",
    "sheetCalcPr", "sheetProtection", "protectedRanges", "scenarios", "autoFilter",
    "sortState", "dataConsolidate", "customSheetViews", "mergeCells", "phoneticPr",
    "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions",
    "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks",
    "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing",
    "drawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts",
    "extLst"
];

/**
 * A worksheet.
 */
class Sheet {
    // /**
    //  * Creates a new instance of Sheet.
    //  * @param {Workbook} workbook - The parent workbook.
    //  * @param {{}} idNode - The sheet ID node (from the parent workbook).
    //  * @param {{}} node - The sheet node.
    //  * @param {{}} [relationshipsNode] - The optional sheet relationships node.
    //  */
    constructor(workbook, idNode, node, relationshipsNode) {
        this._init(workbook, idNode, node, relationshipsNode);
    }

    /* PUBLIC */

    /**
     * Gets a value indicating whether the sheet is the active sheet in the workbook.
     * @returns {boolean} True if active, false otherwise.
     *//**
     * Make the sheet the active sheet in the workkbok.
     * @param {boolean} active - Must be set to `true`. Deactivating directly is not supported. To deactivate, you should activate a different sheet instead.
     * @returns {Sheet} The sheet.
     */
    active() {
        return new ArgHandler('Sheet.active')
            .case(() => {
                return this.workbook().activeSheet() === this;
            })
            .case('boolean', active => {
                if (!active) throw new Error("Deactivating sheet directly not supported. Activate a different sheet instead.");
                this.workbook().activeSheet(this);
                return this;
            })
            .handle(arguments);
    }

    /**
     * Get the active cell in the sheet.
     * @returns {Cell} The active cell.
     *//**
     * Set the active cell in the workbook.
     * @param {string|Cell} cell - The cell or address of cell to activate.
     * @returns {Sheet} The sheet.
     *//**
     * Set the active cell in the workbook by row and column.
     * @param {number} rowNumber - The row number of the cell.
     * @param {string|number} columnNameOrNumber - The column name or number of the cell.
     * @returns {Sheet} The sheet.
     */
    activeCell() {
        const sheetViewNode = this._getOrCreateSheetViewNode();
        let selectionNode = xmlq.findChild(sheetViewNode, "selection");
        return new ArgHandler('Sheet.activeCell')
            .case(() => {
                const cellAddress = selectionNode ? selectionNode.attributes.activeCell : "A1";
                return this.cell(cellAddress);
            })
            .case(['number', '*'], (rowNumber, columnNameOrNumber) => {
                const cell = this.cell(rowNumber, columnNameOrNumber);
                return this.activeCell(cell);
            })
            .case('*', cell => {
                if (!selectionNode) {
                    selectionNode = {
                        name: "selection",
                        attributes: {},
                        children: []
                    };

                    xmlq.appendChild(sheetViewNode, selectionNode);
                }

                if (!(cell instanceof Cell)) cell = this.cell(cell);
                selectionNode.attributes.activeCell = selectionNode.attributes.sqref = cell.address();
                return this;
            })
            .handle(arguments);
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
        const columnNumber = typeof columnNameOrNumber === "string" ? addressConverter.columnNameToNumber(columnNameOrNumber) : columnNameOrNumber;

        // If we're already created a column for this column number, return it.
        if (this._columns[columnNumber]) return this._columns[columnNumber];

        // We need to create a new column, which requires a backing col node. There may already exist a node whose min/max cover our column.
        // First, see if there is an existing col node.
        const existingColNode = this._colNodes[columnNumber];

        let colNode;
        if (existingColNode) {
            // If the existing node covered earlier columns than the new one, we need to have a col node to cover the min up to our new node.
            if (existingColNode.attributes.min < columnNumber) {
                // Clone the node and set the max to the column before our new col.
                const beforeColNode = _.cloneDeep(existingColNode);
                beforeColNode.attributes.max = columnNumber - 1;

                // Update the col nodes cache.
                for (let i = beforeColNode.attributes.min; i <= beforeColNode.attributes.max; i++) {
                    this._colNodes[i] = beforeColNode;
                }
            }

            // Make a clone for the new column. Set the min/max to the column number and cache it.
            colNode = _.cloneDeep(existingColNode);
            colNode.attributes.min = columnNumber;
            colNode.attributes.max = columnNumber;
            this._colNodes[columnNumber] = colNode;

            // If the max of the existing node is greater than the nre one, create a col node for that too.
            if (existingColNode.attributes.max > columnNumber) {
                const afterColNode = _.cloneDeep(existingColNode);
                afterColNode.attributes.min = columnNumber + 1;
                for (let i = afterColNode.attributes.min; i <= afterColNode.attributes.max; i++) {
                    this._colNodes[i] = afterColNode;
                }
            }
        } else {
            // The was no existing node so create a new one.
            colNode = {
                name: 'col',
                attributes: {
                    min: columnNumber,
                    max: columnNumber
                },
                children: []
            };

            this._colNodes[columnNumber] = colNode;
        }

        // Create the new column and cache it.
        const column = new Column(this, colNode);
        this._columns[columnNumber] = column;
        return column;
    }

    /**
     * Gets a defined name scoped to the sheet.
     * @param {string} name - The defined name.
     * @returns {undefined|string|Cell|Range|Row|Column} What the defined name refers to or undefined if not found. Will return the string formula if not a Row, Column, Cell, or Range.
     *//**
     * Set a defined name scoped to the sheet.
     * @param {string} name - The defined name.
     * @param {string|Cell|Range|Row|Column} refersTo - What the name refers to.
     * @returns {Workbook} The workbook.
     */
    definedName() {
        return new ArgHandler("Workbook.definedName")
            .case('string', name => {
                return this.workbook().scopedDefinedName(this, name);
            })
            .case(['string', '*'], (name, refersTo) => {
                this.workbook().scopedDefinedName(this, name, refersTo);
                return this;
            })
            .handle(arguments);
    }

    /**
     * Deletes the sheet and returns the parent workbook.
     * @returns {Workbook} The workbook.
     */
    delete() {
        this.workbook().deleteSheet(this);
        return this.workbook();
    }

    /**
     * Find the given pattern in the sheet and optionally replace it.
     * @param {string|RegExp} pattern - The pattern to look for. Providing a string will result in a case-insensitive substring search. Use a RegExp for more sophisticated searches.
     * @param {string|function} [replacement] - The text to replace or a String.replace callback function. If pattern is a string, all occurrences of the pattern in each cell will be replaced.
     * @returns {Array.<Cell>} The matching cells.
     */
    find(pattern, replacement) {
        pattern = regexify(pattern);

        let matches = [];
        this._rows.forEach(row => {
            if (!row) return;
            matches = matches.concat(row.find(pattern, replacement));
        });

        return matches;
    }

    /**
     * Gets a value indicating whether this sheet's grid lines are visible.
     * @returns {boolean} True if selected, false if not.
     *//**
     * Sets whether this sheet's grid lines are visible.
     * @param {boolean} selected - True to make visible, false to hide.
     * @returns {Sheet} The sheet.
     */
    gridLinesVisible() {
        const sheetViewNode = this._getOrCreateSheetViewNode();
        return new ArgHandler('Sheet.gridLinesVisible')
            .case(() => {
                return sheetViewNode.attributes.showGridLines === 1 || sheetViewNode.attributes.showGridLines === undefined;
            })
            .case('boolean', visible => {
                sheetViewNode.attributes.showGridLines = visible ? 1 : 0;
                return this;
            })
            .handle(arguments);
    }

    /**
     * Gets a value indicating if the sheet is hidden or not.
     * @returns {boolean|string} True if hidden, false if visible, and 'very' if very hidden.
     *//**
     * Set whether the sheet is hidden or not.
     * @param {boolean|string} hidden - True to hide, false to show, and 'very' to make very hidden.
     * @returns {Sheet} The sheet.
     */
    hidden() {
        return new ArgHandler('Sheet.hidden')
            .case(() => {
                if (this._idNode.attributes.state === 'hidden') return true;
                if (this._idNode.attributes.state === 'veryHidden') return "very";
                return false;
            })
            .case('*', hidden => {
                if (hidden) {
                    const visibleSheets = _.filter(this.workbook().sheets(), sheet => !sheet.hidden());
                    if (visibleSheets.length === 1 && visibleSheets[0] === this) {
                        throw new Error("This sheet may not be hidden as a workbook must contain at least one visible sheet.");
                    }

                    // If activate, activate the first other visible sheet.
                    if (this.active()) {
                        const activeIndex = visibleSheets[0] === this ? 1 : 0;
                        visibleSheets[activeIndex].active(true);
                    }
                }

                if (hidden === 'very') this._idNode.attributes.state = 'veryHidden';
                else if (hidden) this._idNode.attributes.state = 'hidden';
                else delete this._idNode.attributes.state;
                return this;
            })
            .handle(arguments);
    }

    /**
     * Move the sheet.
     * @param {number|string|Sheet} [indexOrBeforeSheet] The index to move the sheet to or the sheet (or name of sheet) to move this sheet before. Omit this argument to move to the end of the workbook.
     * @returns {Sheet} The sheet.
     */
    move(indexOrBeforeSheet) {
        this.workbook().moveSheet(this, indexOrBeforeSheet);
        return this;
    }

    /**
     * Get the name of the sheet.
     * @returns {string} The sheet name.
     *//**
     * Set the name of the sheet. *Note: this method does not rename references to the sheet so formulas, etc. can be broken. Use with caution!*
     * @param {string} name - The name to set to the sheet.
     * @returns {Sheet} The sheet.
     */
    name() {
        return new ArgHandler('Sheet.name')
            .case(() => {
                return this._idNode.attributes.name;
            })
            .case('string', name => {
                this._idNode.attributes.name = name;
                return this;
            })
            .handle(arguments);
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
        return new ArgHandler('Sheet.range')
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
     * Get the tab color. (See style [Color](#color).)
     * @returns {undefined|Color} The color or undefined if not set.
     *//**
     * Sets the tab color. (See style [Color](#color).)
     * @returns {Color|string|number} color - Color of the tab. If string, will set an RGB color. If number, will set a theme color.
     */
    tabColor() {
        return new ArgHandler("Sheet.tabColor")
            .case(() => {
                const tabColorNode = xmlq.findChild(this._sheetPrNode, "tabColor");
                if (!tabColorNode) return;

                const color = {};
                if (tabColorNode.attributes.hasOwnProperty('rgb')) color.rgb = tabColorNode.attributes.rgb;
                else if (tabColorNode.attributes.hasOwnProperty('theme')) color.theme = tabColorNode.attributes.theme;
                else if (tabColorNode.attributes.hasOwnProperty('indexed')) color.rgb = colorIndexes[tabColorNode.attributes.indexed];

                if (tabColorNode.attributes.hasOwnProperty('tint')) color.tint = tabColorNode.attributes.tint;

                return color;
            })
            .case("string", rgb => this.tabColor({ rgb }))
            .case("integer", theme => this.tabColor({ theme }))
            .case("nil", () => {
                xmlq.removeChild(this._sheetPrNode, "tabColor");
                return this;
            })
            .case("object", color => {
                const tabColorNode = xmlq.appendChildIfNotFound(this._sheetPrNode, "tabColor");
                xmlq.setAttributes(tabColorNode, {
                    rgb: color.rgb && color.rgb.toUpperCase(),
                    indexed: null,
                    theme: color.theme,
                    tint: color.tint
                });

                return this;
            })
            .handle(arguments);
    }

    /**
     * Gets a value indicating whether this sheet is selected.
     * @returns {boolean} True if selected, false if not.
     *//**
     * Sets whether this sheet is selected.
     * @param {boolean} selected - True to select, false to deselected.
     * @returns {Sheet} The sheet.
     */
    tabSelected() {
        const sheetViewNode = this._getOrCreateSheetViewNode();
        return new ArgHandler('Sheet.tabSelected')
            .case(() => {
                return sheetViewNode.attributes.tabSelected === 1;
            })
            .case('boolean', selected => {
                if (selected) sheetViewNode.attributes.tabSelected = 1;
                else delete sheetViewNode.attributes.tabSelected;
                return this;
            })
            .handle(arguments);
    }

    /**
     * Get the range of cells in the sheet that have contained a value or style at any point. Useful for extracting the entire sheet contents.
     * @returns {Range|undefined} The used range or undefined if no cells in the sheet are used.
     */
    usedRange() {
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
     * @returns {Workbook} The parent workbook.
     */
    workbook() {
        return this._workbook;
    }

    /* INTERNAL */

    /**
     * Clear cells that are using a given shared formula ID.
     * @param {number} sharedFormulaId - The shared formula ID.
     * @returns {undefined}
     * @ignore
     */
    clearCellsUsingSharedFormula(sharedFormulaId) {
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
        // This will work after setting Column.style because Column updates the attributes live.
        const colNode = this._colNodes[columnNumber];
        return colNode && colNode.attributes.style;
    }

    /**
     * Call a callback for each column number that has a node defined for it.
     * @param {Function} callback - The callback.
     * @returns {undefined}
     * @ignore
     */
    forEachExistingColumnNumber(callback) {
        _.forEach(this._colNodes, (node, columnNumber) => {
            if (!node) return;
            callback(columnNumber);
        });
    }

    /**
     * Call a callback for each existing row.
     * @param {Function} callback - The callback.
     * @returns {undefined}
     * @ignore
     */
    forEachExistingRow(callback) {
        _.forEach(this._rows, (row, rowNumber) => {
            if (row) callback(row, rowNumber);
        });

        return this;
    }

    /**
     * Get the hyperlink attached to the cell with the given address.
     * @param {string} address - The address of the hyperlinked cell.
     * @returns {string|undefined} The hyperlink or undefined if not set.
     * @ignore
     *//**
     * Set the hyperlink attached to the cell with the given address.
     * @param {string} address - The address to of the hyperlinked cell.
     * @param {boolean} hyperlink - The hyperlink to set or undefined to clear.
     * @returns {Sheet} The sheet.
     * @ignore
     */
    hyperlink() {
        return new ArgHandler('Sheet.hyperlink')
            .case('string', address => {
                const hyperlinkNode = this._hyperlinks[address];
                if (!hyperlinkNode) return;
                const relationship = this._relationships.findById(hyperlinkNode.attributes['r:id']);
                return relationship && relationship.attributes.Target;
            })
            .case(['string', 'nil'], address => {
                delete this._hyperlinks[address];
                return this;
            })
            .case(['string', 'string'], (address, hyperlink) => {
                const relationship = this._relationships.add("hyperlink", hyperlink, "External");
                this._hyperlinks[address] = {
                    name: 'hyperlink',
                    attributes: { ref: address, 'r:id': relationship.attributes.Id },
                    children: []
                };

                return this;
            })
            .handle(arguments);
    }

    /**
     * Increment and return the max shared formula ID.
     * @returns {number} The new max shared formula ID.
     * @ignore
     */
    incrementMaxSharedFormulaId() {
        return ++this._maxSharedFormulaId;
    }

    /**
     * Get a value indicating whether the cells in the given address are merged.
     * @param {string} address - The address to check.
     * @returns {boolean} True if merged, false if not merged.
     * @ignore
     *//**
     * Merge/unmerge cells by adding/removing a mergeCell entry.
     * @param {string} address - The address to merge.
     * @param {boolean} merged - True to merge, false to unmerge.
     * @returns {Sheet} The sheet.
     * @ignore
     */
    merged() {
        return new ArgHandler('Sheet.merge')
            .case('string', address => {
                return this._mergeCells.hasOwnProperty(address);
            })
            .case(['string', '*'], (address, merge) => {
                if (merge) {
                    this._mergeCells[address] = {
                        name: 'mergeCell',
                        attributes: { ref: address },
                        children: []
                    };
                } else {
                    delete this._mergeCells[address];
                }

                return this;
            })
            .handle(arguments);
    }


    /**
     * Gets a Object or undefined of the cells in the given address.
     * @param {string} address - The address to check.
     * @returns {object|boolean} Object or false if not set
     * @ignore
     *//**
     * Removes dataValidation at the given address
     * @param {string} address - The address to remove.
     * @param {boolean} obj - false to delete.
     * @returns {boolean} true if removed.
     * @ignore
     *//**
     * Add dataValidation to cells at the given address if object or string
     * @param {string} address - The address to set.
     * @param {object|string} obj - Object or String to set
     * @returns {Sheet} The sheet.
     * @ignore
     */
    dataValidation() {
        return new ArgHandler('Sheet.dataValidation')
            .case('string', address => {
                if (this._dataValidations[address]) {
                    return {
                        type: this._dataValidations[address].attributes.type,
                        allowBlank: this._dataValidations[address].attributes.allowBlank,
                        showInputMessage: this._dataValidations[address].attributes.showInputMessage,
                        prompt: this._dataValidations[address].attributes.prompt,
                        promptTitle: this._dataValidations[address].attributes.promptTitle,
                        showErrorMessage: this._dataValidations[address].attributes.showErrorMessage,
                        error: this._dataValidations[address].attributes.error,
                        errorTitle: this._dataValidations[address].attributes.errorTitle,
                        operator: this._dataValidations[address].attributes.operator,
                        formula1: this._dataValidations[address].children[0].children[0],
                        formula2: this._dataValidations[address].children[1] ? this._dataValidations[address].children[1].children[0] : undefined
                    };
                } else {
                    return false;
                }
            })
            .case(['string', 'boolean'], (address, obj) => {
                if (this._dataValidations[address]) {
                    if (obj === false) return delete this._dataValidations[address];
                } else {
                    return false;
                }
            })
            .case(['string', '*'], (address, obj) => {
                if (typeof obj === 'string') {
                    this._dataValidations[address] = {
                        name: 'dataValidation',
                        attributes: {
                            type: 'list',
                            allowBlank: false,
                            showInputMessage: false,
                            prompt: '',
                            promptTitle: '',
                            showErrorMessage: false,
                            error: '',
                            errorTitle: '',
                            operator: '',
                            sqref: address
                        },
                        children: [
                            {
                                name: 'formula1',
                                atrributes: {},
                                children: [obj]
                            },
                            {
                                name: 'formula2',
                                atrributes: {},
                                children: ['']
                            }
                        ]
                    };
                } else if (typeof obj === 'object') {
                    this._dataValidations[address] = {
                        name: 'dataValidation',
                        attributes: {
                            type: obj.type ? obj.type : 'list',
                            allowBlank: obj.allowBlank,
                            showInputMessage: obj.showInputMessage,
                            prompt: obj.prompt,
                            promptTitle: obj.promptTitle,
                            showErrorMessage: obj.showErrorMessage,
                            error: obj.error,
                            errorTitle: obj.errorTitle,
                            operator: obj.operator,
                            sqref: address
                        },
                        children: [
                            {
                                name: 'formula1',
                                atrributes: {},
                                children: [
                                    obj.formula1
                                ]
                            },
                            {
                                name: 'formula2',
                                atrributes: {},
                                children: [
                                    obj.formula2
                                ]
                            }
                        ]
                    };
                }
                return this;
            })
            .handle(arguments);
    }

    /**
     * Convert the sheet to a collection of XML objects.
     * @returns {{}} The XML forms.
     * @ignore
     */
    toXmls() {
        // Shallow clone the node so we don't have to remove these children later if they don't belong.
        const node = _.clone(this._node);
        node.children = node.children.slice();

        // Add the columns if needed.
        this._colsNode.children = _.filter(this._colNodes, (colNode, i) => {
            // Columns should only be present if they have attributes other than min/max.
            return colNode && i === colNode.attributes.min && Object.keys(colNode.attributes).length > 2;
        });
        if (this._colsNode.children.length) {
            xmlq.insertInOrder(node, this._colsNode, nodeOrder);
        }

        // Add the hyperlinks if needed.
        this._hyperlinksNode.children = _.values(this._hyperlinks);
        if (this._hyperlinksNode.children.length) {
            xmlq.insertInOrder(node, this._hyperlinksNode, nodeOrder);
        }

        // Add the merge cells if needed.
        this._mergeCellsNode.children = _.values(this._mergeCells);
        if (this._mergeCellsNode.children.length) {
            xmlq.insertInOrder(node, this._mergeCellsNode, nodeOrder);
        }

        // Add the DataValidation cells if needed.
        this._dataValidationsNode.children = _.values(this._dataValidations);
        if (this._dataValidationsNode.children.length) {
            xmlq.insertInOrder(node, this._dataValidationsNode, nodeOrder);
        }

        return {
            id: this._idNode,
            sheet: node,
            relationships: this._relationships
        };
    }

    /**
     * Update the max shared formula ID to the given value if greater than current.
     * @param {number} sharedFormulaId - The new shared formula ID.
     * @returns {undefined}
     * @ignore
     */
    updateMaxSharedFormulaId(sharedFormulaId) {
        if (sharedFormulaId > this._maxSharedFormulaId) {
            this._maxSharedFormulaId = sharedFormulaId;
        }
    }

    /* PRIVATE */

    /**
     * Get the sheet view node if it exists or create it if it doesn't.
     * @returns {{}} The sheet view node.
     * @private
     */
    _getOrCreateSheetViewNode() {
        let sheetViewsNode = xmlq.findChild(this._node, "sheetViews");
        if (!sheetViewsNode) {
            sheetViewsNode = {
                name: "sheetViews",
                attributes: {},
                children: [{
                    name: "sheetView",
                    attributes: {
                        workbookViewId: 0
                    },
                    children: []
                }]
            };

            xmlq.insertInOrder(this._node, sheetViewsNode, nodeOrder);
        }

        return xmlq.findChild(sheetViewsNode, "sheetView");
    }

    /**
     * Initializes the sheet.
     * @param {Workbook} workbook - The parent workbook.
     * @param {{}} idNode - The sheet ID node (from the parent workbook).
     * @param {{}} node - The sheet node.
     * @param {{}} [relationshipsNode] - The optional sheet relationships node.
     * @returns {undefined}
     * @private
     */
    _init(workbook, idNode, node, relationshipsNode) {
        if (!node) {
            node = {
                name: "worksheet",
                attributes: {
                    xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
                    'xmlns:r': "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                    'xmlns:mc': "http://schemas.openxmlformats.org/markup-compatibility/2006",
                    'mc:Ignorable': "x14ac",
                    'xmlns:x14ac': "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
                },
                children: [{
                    name: "sheetData",
                    attributes: {},
                    children: []
                }]
            };
        }

        this._workbook = workbook;
        this._idNode = idNode;
        this._node = node;
        this._maxSharedFormulaId = -1;
        this._mergeCells = {};
        this._dataValidations = {};
        this._hyperlinks = {};

        // Create the relationships.
        this._relationships = new Relationships(relationshipsNode);

        // Delete the optional dimension node
        xmlq.removeChild(this._node, "dimension");

        // Create the rows.
        this._rows = [];
        this._sheetDataNode = xmlq.findChild(this._node, "sheetData");
        this._sheetDataNode.children.forEach(rowNode => {
            const row = new Row(this, rowNode);
            this._rows[row.rowNumber()] = row;
        });
        this._sheetDataNode.children = this._rows;

        // Create the columns node.
        this._columns = [];
        this._colsNode = xmlq.findChild(this._node, "cols");
        if (this._colsNode) {
            xmlq.removeChild(this._node, this._colsNode);
        } else {
            this._colsNode = { name: 'cols', attributes: {}, children: [] };
        }

        // Cache the col nodes.
        this._colNodes = [];
        _.forEach(this._colsNode.children, colNode => {
            const min = colNode.attributes.min;
            const max = colNode.attributes.max;
            for (let i = min; i <= max; i++) {
                this._colNodes[i] = colNode;
            }
        });

        // Create the sheet properties node.
        this._sheetPrNode = xmlq.findChild(this._node, "sheetPr");
        if (!this._sheetPrNode) {
            this._sheetPrNode = { name: 'sheetPr', attributes: {}, children: [] };
            xmlq.insertInOrder(this._node, this._sheetPrNode, nodeOrder);
        }

        // Create the merge cells.
        this._mergeCellsNode = xmlq.findChild(this._node, "mergeCells");
        if (this._mergeCellsNode) {
            xmlq.removeChild(this._node, this._mergeCellsNode);
        } else {
            this._mergeCellsNode = { name: 'mergeCells', attributes: {}, children: [] };
        }

        const mergeCellNodes = this._mergeCellsNode.children;
        this._mergeCellsNode.children = [];
        mergeCellNodes.forEach(mergeCellNode => {
            this._mergeCells[mergeCellNode.attributes.ref] = mergeCellNode;
        });


        // Create the DataValidations.
        this._dataValidationsNode = xmlq.findChild(this._node, "dataValidations");
        if (this._dataValidationsNode) {
            xmlq.removeChild(this._node, this._dataValidationsNode);
        } else {
            this._dataValidationsNode = { name: 'dataValidations', attributes: {}, children: [] };
        }

        const dataValidationNodes = this._dataValidationsNode.children;
        this._dataValidationsNode.children = [];
        dataValidationNodes.forEach(dataValidationNode => {
            this._dataValidations[dataValidationNode.attributes.sqref] = dataValidationNode;
        });


        // Create the hyperlinks.
        this._hyperlinksNode = xmlq.findChild(this._node, "hyperlinks");
        if (this._hyperlinksNode) {
            xmlq.removeChild(this._node, this._hyperlinksNode);
        } else {
            this._hyperlinksNode = { name: 'hyperlinks', attributes: {}, children: [] };
        }

        const hyperlinkNodes = this._hyperlinksNode.children;
        this._hyperlinksNode.children = [];
        hyperlinkNodes.forEach(hyperlinkNode => {
            this._hyperlinks[hyperlinkNode.attributes.ref] = hyperlinkNode;
        });
    }
}

module.exports = Sheet;
