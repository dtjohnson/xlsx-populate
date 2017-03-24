"use strict";

const _ = require("lodash");
const fs = require("fs");
const JSZip = require('jszip');

const externals = require("./externals");
const regexify = require("./regexify");
const blank = require("./blank")();
const xmlq = require("./xmlq");
const Sheet = require("./Sheet");
const ContentTypes = require("./ContentTypes");
const Relationships = require("./Relationships");
const SharedStrings = require("./SharedStrings");
const StyleSheet = require("./StyleSheet");
const XmlParser = require("./XmlParser");
const XmlBuilder = require("./XmlBuilder");
const ArgHandler = require("./ArgHandler");
const addressConverter = require("./addressConverter");

// Options for adding files to zip. Do not create folders and use a fixed time at epoch.
// The default JSZip behavior uses current time, which causes idential workbooks to be different each time.
const zipFileOpts = {
    date: new Date(0),
    createFolders: false
};

// Initialize the parser and builder.
const xmlParser = new XmlParser();
const xmlBuilder = new XmlBuilder();

// Characters not allowed in sheet names.
const badSheetNameChars = ['\\', '/', '*', '[', ']', ':', '?'];

// Excel limits sheet names to 31 chars.
const maxSheetNameLength = 31;

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
        if (process.browser) throw new Error("Workbook.fromFileAsync is not supported in the browser");
        return new externals.Promise((resolve, reject) => {
            fs.readFile(path, (err, data) => {
                if (err) return reject(err);
                resolve(data);
            });
        }).then(data => Workbook.fromDataAsync(data));
    }

    /**
     * Get the active sheet in the workbook.
     * @returns {Sheet} The active sheet.
     *//**
     * Set the active sheet in the workbook.
     * @param {Sheet|string|number} sheet - The sheet or name of sheet or index of sheet to activate. The sheet must not be hidden.
     * @returns {Workbook} The workbook.
     */
    activeSheet() {
        const bookViewsNode = xmlq.findChild(this._node, "bookViews");
        const workbookViewNode = xmlq.findChild(bookViewsNode, "workbookView");

        return new ArgHandler('Workbook.activeSheet')
            .case(() => {
                const activeTabIndex = workbookViewNode.attributes.activeTab || 0;
                return this._sheets[activeTabIndex];
            })
            .case('*', sheet => {
                // Get the sheet from name/index if needed.
                if (!(sheet instanceof Sheet)) sheet = this.sheet(sheet);

                // Check if the sheet is hidden.
                if (sheet.hidden()) throw new Error("You may not activate a hidden sheet.");

                // Get the index of the sheet.
                const sheetIndex = this._sheets.indexOf(sheet);
                if (sheetIndex < 0) throw new Error('Invalid sheet.');

                // Deselect all sheets except the active one (copying Excel behavior).
                _.forEach(this._sheets, current => {
                    current.tabSelected(current === sheet);
                });

                // Set the active tab attribute to the sheet index.
                workbookViewNode.attributes.activeTab = sheetIndex;

                return this;
            })
            .handle(arguments);
    }

    /**
     * Add a new sheet to the workbook.
     * @param {string} name - The name of the sheet. Must be unique, less than 31 characters, and may not contain the following characters: \ / * [ ] : ?
     * @param {number|string|Sheet} [indexOrBeforeSheet] The index to move the sheet to or the sheet (or name of sheet) to move this sheet before. Omit this argument to move to the end of the workbook.
     * @returns {Sheet} The new sheet.
     */
    addSheet(name, indexOrBeforeSheet) {
        // Validate the sheet name.
        if (!name || typeof name !== "string") throw new Error("Invalid sheet name.");
        if (_.some(badSheetNameChars, char => name.indexOf(char) >= 0)) throw new Error(`Sheet name may not contain any of the following characters: ${badSheetNameChars.join(" ")}`);
        if (name.length > maxSheetNameLength) throw new Error(`Sheet name may not be greater than ${maxSheetNameLength} characters.`);
        if (this.sheet(name)) throw new Error(`Sheet with name "${name}" already exists.`);

        // Capture the current active sheet so we can restore it after moving.
        // (The active sheet is stored by index so adding a new sheet before the active sheet will mess up the active sheet.)
        const currentActiveSheet = this.activeSheet();

        // Get the destination index of new sheet.
        let index;
        if (_.isNil(indexOrBeforeSheet)) {
            index = this._sheets.length;
        } else if (_.isInteger(indexOrBeforeSheet)) {
            index = indexOrBeforeSheet;
        } else {
            if (!(indexOrBeforeSheet instanceof Sheet)) {
                indexOrBeforeSheet = this.sheet(indexOrBeforeSheet);
                if (!indexOrBeforeSheet) throw new Error("Invalid before sheet reference.");
            }

            index = this._sheets.indexOf(indexOrBeforeSheet);
        }

        // Add a new relationship for the new sheet and create the new sheet ID node.
        const relationship = this._relationships.add("worksheet"); // Leave target blank as it will be filled later.
        const sheetIdNode = {
            name: "sheet",
            attributes: {
                name,
                sheetId: ++this._maxSheetId,
                'r:id': relationship.attributes.Id
            },
            children: []
        };

        // Create the new sheet.
        const sheet = new Sheet(this, sheetIdNode);

        // Insert the sheet at the appropriate index.
        this._sheets.splice(index, 0, sheet);

        // Reactivate the previously active sheet.
        this.activeSheet(currentActiveSheet);

        return sheet;
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

        this._sheetsNode.children = [];
        this._sheets.forEach((sheet, i) => {
            const sheetPath = `xl/worksheets/sheet${i + 1}.xml`;
            const sheetRelsPath = `xl/worksheets/_rels/sheet${i + 1}.xml.rels`;
            const sheetObjs = sheet.toObject();
            const relationship = this._relationships.findById(sheetObjs.id.attributes['r:id']);
            relationship.attributes.Target = `worksheets/sheet${i + 1}.xml`;
            this._sheetsNode.children.push(sheetObjs.id);
            this._zip.file(sheetPath, xmlBuilder.build(sheetObjs.sheet), zipFileOpts);
            if (sheetObjs.relationships) {
                this._zip.file(sheetRelsPath, xmlBuilder.build(sheetObjs.relationships), zipFileOpts);
            } else {
                this._zip.remove(sheetRelsPath);
            }
        });

        // Convert the various components to XML strings and add them to the zip.
        this._zip.file("[Content_Types].xml", xmlBuilder.build(this._contentTypes.toObject()), zipFileOpts);
        this._zip.file("xl/_rels/workbook.xml.rels", xmlBuilder.build(this._relationships.toObject()), zipFileOpts);
        this._zip.file("xl/sharedStrings.xml", xmlBuilder.build(this._sharedStrings.toObject()), zipFileOpts);
        this._zip.file("xl/styles.xml", xmlBuilder.build(this._styleSheet.toObject()), zipFileOpts);
        this._zip.file("xl/workbook.xml", xmlBuilder.build(this._node), zipFileOpts);

        // Generate the zip.
        return this._zip.generateAsync({
            type,
            compression: "DEFLATE",
            mimeType: Workbook.MIME_TYPE
        });
    }

    /**
     * Move a sheet to a new position.
     * @param {Sheet|string|number} sheet - The sheet or name of sheet or index of sheet to move.
     * @param {number|string|Sheet} [indexOrBeforeSheet] The index to move the sheet to or the sheet (or name of sheet) to move this sheet before. Omit this argument to move to the end of the workbook.
     * @returns {Workbook} The workbook.
     */
    moveSheet(sheet, indexOrBeforeSheet) {
        // Capture the current active sheet so we can restore it after moving.
        // (The active sheet is stored by index so moving the sheet will mess up the active sheet.)
        const currentActiveSheet = this.activeSheet();

        // Get the sheet to move.
        if (!(sheet instanceof Sheet)) {
            sheet = this.sheet(sheet);
            if (!sheet) throw new Error("Invalid move sheet reference.");
        }

        // Get the to/from indexes.
        const from = this._sheets.indexOf(sheet);
        let to;
        if (_.isNil(indexOrBeforeSheet)) {
            to = this._sheets.length - 1;
        } else if (_.isInteger(indexOrBeforeSheet)) {
            to = indexOrBeforeSheet;
        } else {
            if (!(indexOrBeforeSheet instanceof Sheet)) {
                indexOrBeforeSheet = this.sheet(indexOrBeforeSheet);
                if (!indexOrBeforeSheet) throw new Error("Invalid before sheet reference.");
            }

            to = this._sheets.indexOf(indexOrBeforeSheet);
        }

        // Insert the sheet at the appropriate place.
        this._sheets.splice(to, 0, this._sheets.splice(from, 1)[0]);

        // Reactivate the previously active sheet.
        this.activeSheet(currentActiveSheet);

        return this;
    }

    /**
     * Gets the sheet with the provided name or index (0-based).
     * @param {string|number} sheetNameOrIndex - The sheet name or index.
     * @returns {Sheet|undefined} The sheet or undefined if not found.
     */
    sheet(sheetNameOrIndex) {
        if (_.isInteger(sheetNameOrIndex)) return this._sheets[sheetNameOrIndex];
        return _.find(this._sheets, sheet => sheet.name() === sheetNameOrIndex);
    }

    /**
     * Get an array of all the sheets in the workbook.
     * @returns {Array.<Sheet>} The sheets.
     */
    sheets() {
        return this._sheets.slice(0);
    }

    /**
     * Write the workbook to file. (Not supported in browsers.)
     * @param {string} path - The path of the file to write.
     * @returns {Promise.<undefined>} A promise.
     */
    toFileAsync(path) {
        if (process.browser) throw new Error("Workbook.toFileAsync is not supported in the browser.");
        return this.outputAsync()
            .then(data => new externals.Promise((resolve, reject) => {
                fs.writeFile(path, data, err => {
                    if (err) return reject(err);
                    resolve();
                });
            }));
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
        this._maxSheetId = 0;
        this._sheets = [];

        return JSZip.loadAsync(data)
            .then(zip => {
                this._zip = zip;
                return this._parseNodesAsync([
                    "[Content_Types].xml",
                    "xl/_rels/workbook.xml.rels",
                    "xl/sharedStrings.xml",
                    "xl/styles.xml",
                    "xl/workbook.xml"
                ]);
            })
            .then(nodes => {
                const contentTypesNode = nodes[0];
                const relationshipsNode = nodes[1];
                const sharedStringsNode = nodes[2];
                const styleSheetNode = nodes[3];
                const workbookNode = nodes[4];

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
                return externals.Promise.all(_.map(this._sheetsNode.children, (sheetIdNode, i) => {
                    if (sheetIdNode.attributes.sheetId > this._maxSheetId) this._maxSheetId = sheetIdNode.attributes.sheetId;

                    return this._parseNodesAsync([`xl/worksheets/sheet${i + 1}.xml`, `xl/worksheets/_rels/sheet${i + 1}.xml.rels`])
                        .then(nodes => {
                            const sheetNode = nodes[0];
                            const sheetRelationshipsNode = nodes[1];

                            // Insert at position i as the promises will resolve at different times.
                            this._sheets[i] = new Sheet(this, sheetIdNode, sheetNode, sheetRelationshipsNode);
                        });
                }));
            })
            .then(() => this);
    }

    /**
     * Parse files out of zip into XML node objects.
     * @param {Array.<string>} names - The file names to parse.
     * @returns {Promise.<Array.<{}>>} An array of the parsed objects.
     * @private
     */
    _parseNodesAsync(names) {
        return externals.Promise.all(_.map(names, name => this._zip.file(name)))
            .then(files => externals.Promise.all(_.map(files, file => file && file.async("string"))))
            .then(texts => externals.Promise.all(_.map(texts, text => text && xmlParser.parseAsync(text))));
    }
}

/**
 * The XLSX mime type.
 * @type {string}
 * @ignore
 */
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
