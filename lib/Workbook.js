"use strict";

const Promise = require("bluebird");
const fs = Promise.promisifyAll(require("fs"));
const _ = require("lodash");
var JSZip = require('jszip');
var utils = require('./utils');
var Sheet = require('./Sheet');
const SharedStringsTable = require('./SharedStringsTable');
var path = require("path");
var DOMParser = require('xmldom').DOMParser;
var parser = new DOMParser();
var xpath = require("./xpath");

JSZip.external.Promise = Promise;

class Workbook {
    /**
     * Initialize the workbook. (This is separated from the constructor to ease testing.)
     * @param {Buffer} data - File buffer of the Excel workbook.
     * @returns {undefined}
     * @private
     */
    _initializeAsync(data) {
        return JSZip.loadAsync(data)
            .then(zip => {
                this._zip = zip;

                const sharedStringsFile = this._zip.file("xl/sharedStrings.xml");
                return Promise.all([
                    this._zip.file("xl/workbook.xml").async("string"),
                    this._zip.file("xl/_rels/workbook.xml.rels").async("string"),
                    sharedStringsFile && sharedStringsFile.async("string")
                ]);
            })
            .spread((workbookText, relsText, sharedStringsText) => {
                this._workbookXML = parser.parseFromString(workbookText).documentElement;
                this._relsXML = parser.parseFromString(relsText).documentElement;

                const sharedStringsXML = sharedStringsText && parser.parseFromString(sharedStringsText).documentElement;
                this._sharedStringsTable = new SharedStringsTable(sharedStringsXML);

                this._sheets = [];
                this._sheetsNode = xpath("sml:sheets", this._workbookXML)[0];

                var sheetNodes = this._sheetsNode.childNodes;
                return Promise.map(_.range(this._sheetsNode.childNodes.length), i => {
                    return this._zip.file("xl/worksheets/sheet" + (i + 1) + ".xml").async("string")
                        .then(sheetText => {
                            var sheetXML = parser.parseFromString(sheetText).documentElement;

                            // This is a blunt way to make sure formula values get updated.
                            // It just clears all stored values in case the referenced cell values change.
                            var valueNodes = xpath("sml:sheetData/sml:row/sml:c/sml:f/../sml:*[name(.) !='f']", sheetXML);
                            valueNodes.forEach(function (valueNode) {
                                valueNode.parentNode.removeChild(valueNode);
                            });

                            var sheet = new Sheet(this, sheetNodes[i], sheetXML);
                            this._sheets.push(sheet);
                        });
                });
            })
            .return(this);
    }

    /**
     * Create a new sheet.
     * @param {string} sheetName - The name of the sheet. Must be unique.
     * @param {number} [index] - The position of the sheet (0-based). Omit to insert at the end.
     * @returns {Sheet} The new sheet.
     */
    createSheet(sheetName, index) {
        if (index === undefined) index = this._sheets.length;
        if (!utils.isInteger(index) || index < 0 || index > this._sheets.length) {
            throw new Error("Invalid sheet index.");
        }

        // Create the new XML nodes.
        var sheetXML = parser.parseFromString('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>').documentElement;
        var sheetNode = parser.parseFromString('<sheet name="' + sheetName + '"/>').documentElement;

        // Insert the sheet definition node in the right place.
        if (index === this._sheets.length) {
            this._sheetsNode.appendChild(sheetNode);
        } else {
            this._sheetsNode.insertBefore(sheetNode, this._sheetsNode.childNodes[index]);
        }

        // Clear all the old sheet rel nodes.
        for (var i = this._relsXML.childNodes.length - 1; i >= 0; i--) {
            var rnode = this._relsXML.childNodes[i];
            if (rnode.getAttribute("Type") === "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet") {
                this._relsXML.removeChild(rnode);
            }
        }

        // Fix the sheet IDs to match the order.
        for (var j = 0; j < this._sheetsNode.childNodes.length; j++) {
            var id = j + 1;
            var snode = this._sheetsNode.childNodes[j];
            snode.setAttribute("sheetId", id);
            snode.setAttribute("r:id", "xpopId" + id);

            // Create a new sheet rel node.
            var relNode = parser.parseFromString('<Relationship Id="xpopId' + id + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + id + '.xml"/>');
            this._relsXML.appendChild(relNode);
        }

        // Create the sheet and store it.
        var sheet = new Sheet(this, sheetNode, sheetXML);
        this._sheets.splice(index, 0, sheet);
        return sheet;
    }

    /**
     * Gets the sheet with the provided name or index (0-based).
     * @param {string|number} sheetNameOrIndex - The sheet name or index.
     * @returns {Sheet} The sheet, if found.
     */
    sheet(sheetNameOrIndex) {
        if (utils.isInteger(sheetNameOrIndex)) return this._sheets[sheetNameOrIndex];

        for (var i = 0; i < this._sheets.length; i++) {
            var sheet = this._sheets[i];
            if (sheet.name() === sheetNameOrIndex) return sheet;
        }
    }

    /**
     * Get a named cell. (Assumes names with workbook scope pointing to single cells.)
     * @param {string} cellName - The name of the cell.
     * @returns {Cell} The cell, if found.
     */
    namedCell(cellName) {
        var definedName = xpath("sml:definedNames/sml:definedName[@name='" + cellName + "']", this._workbookXML)[0];
        if (!definedName) return;

        var address = definedName.firstChild.nodeValue;
        var ref = utils.addressToRowAndColumn(address);
        if (!ref) return;

        return this.sheet(ref.sheet).cell(ref.row, ref.column);
    }

    /**
     * Gets the output.
     * @returns {Buffer} A node buffer for the generated Excel workbook.
     */
    outputAsync() {
        this._zip.file("xl/workbook.xml", this._workbookXML.toString());
        this._zip.file("xl/_rels/workbook.xml.rels", this._relsXML.toString());
        this._zip.file("xl/sharedStrings.xml", this._sharedStringsTable.toString());

        for (var i = 0; i < this._sheets.length; i++) {
            var index = i + 1;
            var sheet = this._sheets[i];
            this._zip.file("xl/worksheets/sheet" + index + ".xml", sheet._sheetXML.toString());
        }

        // Kill the calc chain since this will corrupt the file is formulas are removed.
        this._zip.remove("xl/calcChain.xml");

        return this._zip.generateAsync({ type: "nodebuffer" });
    }

    toFileAsync(path) {
        return this.outputAsync()
            .then(data => fs.writeFileAsync(path, data));
    }

    static fromDataAsync(data) {
        return new Workbook()._initializeAsync(data);
    };

    static fromFileAsync(path) {
        return fs.readFileAsync(path)
            .then(data => Workbook.fromDataAsync(data));
    };

    static fromBlankAsync() {
        return Workbook.fromFileAsync(path.join(__dirname, "blank.xlsx"));
    };
}

module.exports = Workbook;
