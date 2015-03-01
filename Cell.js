"use strict";

var etree = require('elementtree'),
    subelement = etree.SubElement,
    utils = require('./utils');

/**
 * Initializes a new Cell.
 * @param {Sheet} sheet
 * @param {number} row
 * @param {number} column
 * @param {etree.SubElement} cellNode
 * @constructor
 */
var Cell = function (sheet, row, column, cellNode) {
    this._sheet = sheet;
    this._row = row;
    this._column = column;
    this._cellNode = cellNode;
};

/**
 * Gets the parent sheet.
 * @returns {Sheet}
 */
Cell.prototype.getSheet = function () {
    return this._sheet;
};

/**
 * Gets the row of the cell.
 * @returns {number}
 */
Cell.prototype.getRow = function () {
    return this._row;
};

/**
 * Gets the column of the cell.
 * @returns {number}
 */
Cell.prototype.getColumn = function () {
    return this._column;
};

/**
 * Gets the address of the cell (e.g. "A5").
 * @returns {string}
 */
Cell.prototype.getAddress = function () {
    return utils.rowAndColumnToAddress(this._row, this._column);
};

/**
 * Gets the full address of the cell including sheet (e.g. "Sheet1!A5").
 * @returns {string}
 */
Cell.prototype.getFullAddress = function () {
    return utils.rowAndColumnToAddress(this._row, this._column, this.getSheet().getName());
};

/**
 * Sets the value of the cell.
 * @param {*} value
 * @returns {Cell}
 */
Cell.prototype.setValue = function (value) {
    this._clearContents();

    if (typeof value === "string") {
        this._cellNode.attrib.t = "inlineStr";
        var isNode = subelement(this._cellNode, "is");
        var tNode = subelement(isNode, "t");
        tNode.text = value;
    } else {
        var vNode = subelement(this._cellNode, "v");
        vNode.text = value;
    }

    return this;
};

/**
 * Sets the formula for a cell (with optional precalculated value).
 * @param {string} formula
 * @param {*=} calculatedValue
 * @returns {Cell}
 */
Cell.prototype.setFormula = function (formula, calculatedValue) {
    this._clearContents();

    var fNode = subelement(this._cellNode, "f");
    fNode.text = formula;

    if (arguments.length > 1) {
        var vNode = subelement(this._cellNode, "v");
        vNode.text = calculatedValue;
    }

    return this;
};

/**
 * Clears the contents from the cell.
 * @private
 */
Cell.prototype._clearContents = function () {
    var self = this;
    this._cellNode.getchildren().forEach(function (childNode) {
        self._cellNode.remove(childNode);
    });

    delete this._cellNode.attrib.t;
};

module.exports = Cell;