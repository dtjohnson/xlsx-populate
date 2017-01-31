"use strict";

// TODO: JSDoc
// TODO: Tests

// TODO: Enum of standard Excel formats.

const Promise = require("bluebird");
const xml2js = Promise.promisifyAll(require("xml2js"));
const fs = require("fs"); // Don't use Bluebird's promisify on fs as it doesn't play nicely with browserify
const JSZip = require('jszip');
const path = require("path");
const utils = require("./utils");
const blank = require("./blank");

const Sheet = require("./Sheet");
const _ContentTypes = require("./_ContentTypes");
const _Relationships = require("./_Relationships");
const _SharedStrings = require("./_SharedStrings");
const _StyleSheet = require("./_StyleSheet");

JSZip.external.Promise = Promise;

const DOMParser = require('xmldom').DOMParser;
const parser = new DOMParser();

class Workbook {
    /**
     * Initialize the workbook. (This is separated from the constructor to ease testing.)
     * @param {Buffer} data - File buffer of the Excel workbook.
     * @returns {undefined}
     * @private
     */
    _initializeAsync(data) {
        this._sheets = [];
        return JSZip.loadAsync(data)
            .then(zip => {
                this._zip = zip;

                const sharedStringsFile = this._zip.file("xl/sharedStrings.xml");
                return Promise.all([
                    this._zip.file("[Content_Types].xml").async("string")
                        .then(text => xml2js.parseStringAsync(text)),
                    this._zip.file("xl/_rels/workbook.xml.rels").async("string"),
                    sharedStringsFile && sharedStringsFile.async("string"),
                    this._zip.file("xl/styles.xml").async("string"),
                    this._zip.file("xl/workbook.xml").async("string")
                ]);
            })
            .spread((contentTypesNode, relationshipsText, sharedStringsText, styleSheetText, workbookText) => {
                this._contentTypes = new _ContentTypes(contentTypesNode);
                this._relationships = new _Relationships(relationshipsText);
                this._sharedStrings = new _SharedStrings(sharedStringsText);
                this._styleSheet = new _StyleSheet(styleSheetText);
                this._xml = parser.parseFromString(workbookText);

                if (!this._relationships.findByType("sharedStrings")) {
                    this._relationships.add("sharedStrings", "sharedStrings.xml");
                }

                if (!this._contentTypes.findByPartName("/xl/sharedStrings.xml")) {
                    this._contentTypes.add("/xl/sharedStrings.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
                }

                // Kill the calc chain. It's not required and the workbook will corrupt unless we keep it up to date.
                this._zip.remove("xl/calcChain.xml");

                this._sheetsNode = this._xml.documentElement.getElementsByTagName("sheets")[0];
                const sheetNodes = utils.mapChildElements(this._sheetsNode, node => node);
                return Promise.map(sheetNodes, (sheetNode, i) => {
                    return this._zip.file(`xl/worksheets/sheet${i + 1}.xml`).async("string")
                        .then(sheetText => this._sheets.push(new Sheet(this, sheetNode, sheetText)));
                });
            })
            .return(this);
    }

    createSheet() {
        // TODO
    }

    // TODO: Kill these named Cell/Range/Group in favor of a single select method?
    namedCell() {

    }

    namedRange() {

    }

    find(pattern) {
        pattern = utils.getRegExpForSearch(pattern);

        let matches = [];
        this._sheets.forEach(sheet => {
            matches = matches.concat(sheet.find(pattern));
        });

        return matches;
    }

    replace(pattern, replacement) {
        pattern = utils.getRegExpForSearch(pattern);

        let count = 0;
        this._sheets.forEach(sheet => {
            count += sheet.replace(pattern, replacement);
        });

        return count;
    }

    /**
     * Gets the output.
     * @returns {Buffer} A node buffer for the generated Excel workbook.
     */
    outputAsync(type) {
        if (!type) {
            type = process.browser ? "blob" : "nodebuffer";
        }

        const builder = new xml2js.Builder();
        this._zip.file("[Content_Types].xml", builder.buildObject(this._contentTypes.toObject()));
        this._zip.file("xl/_rels/workbook.xml.rels", this._relationships.toString());
        this._zip.file("xl/sharedStrings.xml", this._sharedStrings.toString());
        this._zip.file("xl/styles.xml", this._styleSheet.toString());
        this._zip.file("xl/workbook.xml", this._xml.toString());

        this._sheets.forEach((sheet, i) => {
            this._zip.file(`xl/worksheets/sheet${i + 1}.xml`, sheet.toString());
        });

        return this._zip.generateAsync({
            type,
            compression: "DEFLATE",
            mimeType: Workbook.MIME_TYPE
        });
    }

    /**
     * Gets the sheet with the provided name or index (0-based).
     * @param {string|number} sheetNameOrIndex - The sheet name or index.
     * @returns {Sheet} The sheet, if found.
     */
    sheet(sheetNameOrIndex) {
        if (Number.isInteger(sheetNameOrIndex)) return this._sheets[sheetNameOrIndex];
        return this._sheets.find(sheet => sheet.name() === sheetNameOrIndex);
    }

    toFileAsync(path) {
        if (process.browser) throw new Error("Not supported");
        return this.outputAsync()
            .then(data => Promise.fromCallback(cb => fs.writeFile(path, data, cb)));
    }

    group() {
        // TODO
    }

    static fromDataAsync(data) {
        return new Workbook()._initializeAsync(data);
    }

    static fromFileAsync(path) {
        if (process.browser) throw new Error("Not supported");
        return Promise.fromCallback(cb => fs.readFile(path, cb))
            .then(data => Workbook.fromDataAsync(data));
    }

    static fromBlankAsync() {
        return Workbook.fromDataAsync(blank);
    }
}

Workbook.MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

module.exports = Workbook;

/*
xl/workbook.xml

<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
	<fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="16925"/>
	<workbookPr defaultThemeVersion="164011"/>
	<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
		<mc:Choice Requires="x15">
			<x15ac:absPath url="\path\to\file" xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac"/>
		</mc:Choice>
	</mc:AlternateContent>
	<bookViews>
		<workbookView xWindow="3720" yWindow="0" windowWidth="27870" windowHeight="12795"/>
	</bookViews>
	<sheets>
		<sheet name="Sheet1" sheetId="1" r:id="rId1"/>
	</sheets>
	<calcPr calcId="171027"/>
	<extLst>
		<ext uri="{140A7094-0E35-4892-8432-C4D2E57EDEB5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
			<x15:workbookPr chartTrackingRefBase="1"/>
		</ext>
	</extLst>
</workbook>
// */
