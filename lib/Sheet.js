"use strict";

var utils = require('./utils'),
    etree = require('elementtree'),
    subelement = etree.SubElement,
    Cell = require('./Cell');

/**
 * Initializes a new Sheet.
 * @param {Workbook} workbook
 * @param {string} name
 * @param {etree.Element} sheetNode - The node defining the sheet in the workbook.xml.
 * @param {etree.Element} sheetXML
 * @constructor
 */
var Sheet = function (workbook, sheetNode, sheetXML) {
    this._workbook = workbook;
    this._sheetNode = sheetNode;
    this._sheetXML = sheetXML;
};

/**
 * Gets the parent workbook.
 * @returns {Workbook}
 */
Sheet.prototype.getWorkbook = function () {
    return this._workbook;
};

/**
 * Gets the name of the sheet.
 * @returns {string}
 */
Sheet.prototype.getName = function () {
    return this._sheetNode.attrib.name;
};

Sheet.prototype.setName = function (name) {
    this._sheetNode.attrib.name = name;
};

/**
 * Gets the cell with either the provided row and column or address.
 * @returns {Cell}
 */
Sheet.prototype.getCell = function () {
    var row, column, address;
    if (arguments.length === 1) {
        address = arguments[0];
        var ref = utils.addressToRowAndColumn(address);
        row = ref.row;
        column = ref.column;
    } else {
        row = arguments[0];
        column = arguments[1];
        address = utils.rowAndColumnToAddress(row, column);
    }

    var sheetDataNode = this._sheetXML.find("sheetData");
    var rowNode = sheetDataNode.find("row[@r='" + row + "']");
    if (!rowNode) {
        rowNode = subelement(sheetDataNode, "row");
        rowNode.attrib.r = row;
    }

    var cellNode = rowNode.find("c[@r='" + address + "']");
    if (!cellNode) {
        cellNode = subelement(rowNode, "c");
        cellNode.attrib.r = address;
    }

    return new Cell(this, row, column, cellNode);
};

module.exports = Sheet;
