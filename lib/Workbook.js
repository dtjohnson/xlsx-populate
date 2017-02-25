"use strict";

const _ = require("lodash");
const Promise = require("bluebird");
const fs = require("fs"); // Don't use Bluebird's promisify on fs as it doesn't play nicely with browserify
const JSZip = require('jszip');
JSZip.external.Promise = Promise; // Set the JSZip promise to Bluebird so we can use, map, spread, etc.

const regexify = require("./regexify");
const blank = require("./blank");
const xmlq = require("./xmlq");
const Sheet = require("./Sheet");
const ContentTypes = require("./ContentTypes");
const Relationships = require("./Relationships");
const SharedStrings = require("./SharedStrings");
const StyleSheet = require("./StyleSheet");
const XmlParser = require("./XmlParser");
const XmlBuilder = require("./XmlBuilder");
const addressConverter = require("./addressConverter");
const XlsxPopulate = require("./XlsxPopulate");

// Options for adding files to zip. Do not create folders and use a fixed time at epoch.
// The default JSZip behavior uses current time, which causes idential workbooks to be different each time.
const zipFileOpts = {
    date: new Date(0),
    createFolders: false
};

// Initialize the parser and builder.
const xmlParser = new XmlParser();
const xmlBuilder = new XmlBuilder();

/**
 * A workbook.
 */
class Workbook {
    /**
     * Create a new blank workbook.
     * @returns {Promise.<Workbook>} The workbook.
     * @ignore
     */
    static fromBlankAsync() {
        return Workbook.fromDataAsync(blank);
    }

    /**
     * Loads a workbook from a data object. (Supports any supported [JSZip data types]{@link https://stuk.github.io/jszip/documentation/api_jszip/load_async.html}.)
     * @param {string|Array.<number>|ArrayBuffer|Uint8Array|Buffer|Blob|Promise.<*>} data - The data to load.
     * @returns {Promise.<Workbook>} The workbook.
     * @ignore
     */
    static fromDataAsync(data) {
        return new Workbook()._initAsync(data);
    }

    /**
     * Loads a workbook from file.
     * @param {string} path - The path to the workbook.
     * @returns {Promise.<Workbook>} The workbook.
     * @ignore
     */
    static fromFileAsync(path) {
        if (process.browser) throw new Error("Not supported");
        return Promise.fromCallback(cb => fs.readFile(path, cb))
            .then(data => Workbook.fromDataAsync(data));
    }

    /**
     * Gets a defined name scoped to the workbook.
     * @param {string} name - The defined name.
     * @returns {undefined|Cell|Range|Row|Column} The named selection or undefined if name not found.
     * @throws {Error} Will throw if address in defined name is not supported.
     */
    definedName(name) {
        return this.scopedDefinedName(name);
    }

    /**
     * Find the given pattern in the workbook and optionally replace it.
     * @param {string|RegExp} pattern - The pattern to look for. Providing a string will result in a case-insensitive substring search. Use a RegExp for more sophisticated searches.
     * @param {string|function} [replacement] - The text to replace or a String.replace callback function. If pattern is a string, all occurrences of the pattern in each cell will be replaced.
     * @returns {boolean} A flag indicating if the pattern was found.
     */
    find(pattern, replacement) {
        pattern = regexify(pattern);

        let matches = [];
        this._sheets.forEach(sheet => {
            matches = matches.concat(sheet.find(pattern, replacement));
        });

        return matches;
    }

    /**
     * Generates the workbook output.
     * @param {string} [type] - The type of the data to return. (Supports any supported [JSZip data types]{@link https://stuk.github.io/jszip/documentation/api_jszip/generate_async.html}: base64, binarystring, uint8array, arraybuffer, blob, nodebuffer) Defaults to 'nodebuffer' in Node.js and 'blob' in browsers.
     * @returns {string|Uint8Array|ArrayBuffer|Blob|Buffer} The data.
     */
    outputAsync(type) {
        if (!type) type = process.browser ? "blob" : "nodebuffer";

        // Convert the various components to XML strings and add them to the zip.
        this._zip.file("[Content_Types].xml", xmlBuilder.build(this._contentTypes.toObject()), zipFileOpts);
        this._zip.file("xl/_rels/workbook.xml.rels", xmlBuilder.build(this._relationships.toObject()), zipFileOpts);
        this._zip.file("xl/sharedStrings.xml", xmlBuilder.build(this._sharedStrings.toObject()), zipFileOpts);
        this._zip.file("xl/styles.xml", xmlBuilder.build(this._styleSheet.toObject()), zipFileOpts);
        this._zip.file("xl/workbook.xml", xmlBuilder.build(this._node), zipFileOpts);

        this._sheets.forEach((sheet, i) => {
            this._zip.file(`xl/worksheets/sheet${i + 1}.xml`, xmlBuilder.build(sheet.toObject()), zipFileOpts);
        });

        // Generate the zip.
        return this._zip.generateAsync({
            type,
            compression: "DEFLATE",
            mimeType: XlsxPopulate.MIME_TYPE
        });
    }

    /**
     * Gets the sheet with the provided name or index (0-based).
     * @param {string|number} sheetNameOrIndex - The sheet name or index.
     * @returns {Sheet|undefined} The sheet or undefined if not found.
     */
    sheet(sheetNameOrIndex) {
        if (Number.isInteger(sheetNameOrIndex)) return this._sheets[sheetNameOrIndex];
        return this._sheets.find(sheet => sheet.name() === sheetNameOrIndex);
    }

    /**
     * Write the workbook to file. (Not supported in browsers.)
     * @param {string} path - The path of the file to write.
     * @returns {Promise.<undefined>} A promise.
     */
    toFileAsync(path) {
        if (process.browser) throw new Error("Workbook.toFileAsync is not supported in the browser.");
        return this.outputAsync()
            .then(data => Promise.fromCallback(cb => fs.writeFile(path, data, cb)));
    }

    /**
     * Gets a scoped defined name.
     * @param {string} name - The defined name.
     * @param {Sheet} [sheetScope] - The sheet the name is scoped to, if applicable.
     * @returns {undefined|Cell|Range|Row|Column} The named selection or undefined if name not found.
     * @throws {Error} Will throw if address in defined name is not supported.
     * @ignore
     */
    scopedDefinedName(name, sheetScope) {
        let localSheetId;
        if (sheetScope) localSheetId = this._sheets.indexOf(sheetScope);

        // Get the address from the definedNames node.
        const definedNamesNode = xmlq.findChild(this._node, "definedNames");
        const definedName = definedNamesNode && _.find(definedNamesNode.children, node => node.attributes.name === name && node.attributes.localSheetId === localSheetId);
        const address = definedName && definedName.children[0];
        if (!address) return undefined;

        // Try to parse the address.
        let ref;
        try {
            ref = addressConverter.fromAddress(address);
        } catch (err) {
            throw new Error("Defined name found but value is not currently supported.");
        }

        // Load the appropriate selection type.
        const sheet = this.sheet(ref.sheetName);
        if (ref.type === 'cell') return sheet.cell(ref.rowNumber, ref.columnNumber);
        if (ref.type === 'range') return sheet.range(ref.startRowNumber, ref.startColumnNumber, ref.endRowNumber, ref.endColumnNumber);
        if (ref.type === 'row') return sheet.row(ref.rowNumber);
        if (ref.type === 'column') return sheet.column(ref.columnNumber);

        throw new Error(`Defined name found but value type '${ref.type}' is not currently supported.`);
    }

    /**
     * Get the shared strings table.
     * @returns {SharedStrings} The shared strings table.
     * @ignore
     */
    sharedStrings() {
        return this._sharedStrings;
    }

    /**
     * Get the style sheet.
     * @returns {StyleSheet} The style sheet.
     * @ignore
     */
    styleSheet() {
        return this._styleSheet;
    }

    /**
     * Initialize the workbook. (This is separated from the constructor to ease testing.)
     * @param {string|Array.<number>|ArrayBuffer|Uint8Array|Buffer|Blob|Promise.<*>} data - The data to load.
     * @returns {Promise.<Workbook>} The workbook.
     * @private
     */
    _initAsync(data) {
        this._sheets = [];
        return JSZip.loadAsync(data)
            .then(zip => {
                this._zip = zip;
                return [
                    "[Content_Types].xml",
                    "xl/_rels/workbook.xml.rels",
                    "xl/sharedStrings.xml",
                    "xl/styles.xml",
                    "xl/workbook.xml"
                ];
            })
            .map(name => this._zip.file(name))
            .map(file => file && file.async("string"))
            .map(text => text && xmlParser.parseAsync(text))
            .spread((contentTypesNode, relationshipsNode, sharedStringsNode, styleSheetNode, workbookNode) => {
                // Load the various components.
                this._contentTypes = new ContentTypes(contentTypesNode);
                this._relationships = new Relationships(relationshipsNode);
                this._sharedStrings = new SharedStrings(sharedStringsNode);
                this._styleSheet = new StyleSheet(styleSheetNode);
                this._node = workbookNode;

                // Add the shared strings relationship if it doesn't exist.
                if (!this._relationships.findByType("sharedStrings")) {
                    this._relationships.add("sharedStrings", "sharedStrings.xml");
                }

                // Add the shared string content type if it doesn't exist.
                if (!this._contentTypes.findByPartName("/xl/sharedStrings.xml")) {
                    this._contentTypes.add("/xl/sharedStrings.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
                }

                // Kill the calc chain. It's not required and the workbook will corrupt unless we keep it up to date.
                this._zip.remove("xl/calcChain.xml");

                // Load each sheet.
                this._sheetsNode = xmlq.findChild(this._node, "sheets");
                return Promise.map(this._sheetsNode.children, (sheetIdNode, i) => {
                    return this._zip.file(`xl/worksheets/sheet${i + 1}.xml`)
                        .async("string")
                        .then(text => xmlParser.parseAsync(text))
                        .then(sheetNode => this._sheets.push(new Sheet(this, sheetIdNode, sheetNode)));
                });
            })
            .return(this);
    }
}

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
