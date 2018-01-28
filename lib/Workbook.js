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
const AppProperties = require("./AppProperties");
const Relationships = require("./Relationships");
const SharedStrings = require("./SharedStrings");
const StyleSheet = require("./StyleSheet");
const Encryptor = require("./Encryptor");
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

// Initialize the encryptor if present (can be excluded in browser build).
const encryptor = typeof Encryptor === "function" && new Encryptor();

// Characters not allowed in sheet names.
const badSheetNameChars = ['\\', '/', '*', '[', ']', ':', '?'];

// Excel limits sheet names to 31 chars.
const maxSheetNameLength = 31;

// Order of the nodes as defined by the spec.
const nodeOrder = [
    "fileVersion", "fileSharing", "workbookPr", "workbookProtection", "bookViews", "sheets", "functionGroups",
    "externalReferences", "definedNames", "calcPr", "oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr",
    "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"
];

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
     * @param {{}} [opts] - Options
     * @returns {Promise.<Workbook>} The workbook.
     * @ignore
     */
    static fromDataAsync(data, opts) {
        return new Workbook()._initAsync(data, opts);
    }

    /**
     * Loads a workbook from file.
     * @param {string} path - The path to the workbook.
     * @param {{}} [opts] - Options
     * @returns {Promise.<Workbook>} The workbook.
     * @ignore
     */
    static fromFileAsync(path, opts) {
        if (process.browser) throw new Error("Workbook.fromFileAsync is not supported in the browser");
        return new externals.Promise((resolve, reject) => {
            fs.readFile(path, (err, data) => {
                if (err) return reject(err);
                resolve(data);
            });
        }).then(data => Workbook.fromDataAsync(data, opts));
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
        return new ArgHandler('Workbook.activeSheet')
            .case(() => {
                return this._activeSheet;
            })
            .case('*', sheet => {
                // Get the sheet from name/index if needed.
                if (!(sheet instanceof Sheet)) sheet = this.sheet(sheet);

                // Check if the sheet is hidden.
                if (sheet.hidden()) throw new Error("You may not activate a hidden sheet.");

                // Deselect all sheets except the active one (mirroring ying Excel behavior).
                _.forEach(this._sheets, current => {
                    current.tabSelected(current === sheet);
                });

                this._activeSheet = sheet;

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

        return sheet;
    }

    /**
     * Gets a defined name scoped to the workbook.
     * @param {string} name - The defined name.
     * @returns {undefined|string|Cell|Range|Row|Column} What the defined name refers to or undefined if not found. Will return the string formula if not a Row, Column, Cell, or Range.
     *//**
     * Set a defined name scoped to the workbook.
     * @param {string} name - The defined name.
     * @param {string|Cell|Range|Row|Column} refersTo - What the name refers to.
     * @returns {Workbook} The workbook.
     */
    definedName() {
        return new ArgHandler("Workbook.definedName")
            .case('string', name => {
                return this.scopedDefinedName(undefined, name);
            })
            .case(['string', '*'], (name, refersTo) => {
                this.scopedDefinedName(undefined, name, refersTo);
                return this;
            })
            .handle(arguments);
    }

    /**
     * Delete a sheet from the workbook.
     * @param {Sheet|string|number} sheet - The sheet or name of sheet or index of sheet to move.
     * @returns {Workbook} The workbook.
     */
    deleteSheet(sheet) {
        // Get the sheet to move.
        if (!(sheet instanceof Sheet)) {
            sheet = this.sheet(sheet);
            if (!sheet) throw new Error("Invalid move sheet reference.");
        }

        // Make sure we are not deleting the only visible sheet.
        const visibleSheets = _.filter(this._sheets, sheet => !sheet.hidden());
        if (visibleSheets.length === 1 && visibleSheets[0] === sheet) {
            throw new Error("This sheet may not be deleted as a workbook must contain at least one visible sheet.");
        }

        // Remove the sheet.
        let index = this._sheets.indexOf(sheet);
        this._sheets.splice(index, 1);

        // Set the new active sheet.
        if (sheet === this.activeSheet()) {
            if (index >= this._sheets.length) index--;
            this.activeSheet(index);
        }

        return this;
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
     * Move a sheet to a new position.
     * @param {Sheet|string|number} sheet - The sheet or name of sheet or index of sheet to move.
     * @param {number|string|Sheet} [indexOrBeforeSheet] The index to move the sheet to or the sheet (or name of sheet) to move this sheet before. Omit this argument to move to the end of the workbook.
     * @returns {Workbook} The workbook.
     */
    moveSheet(sheet, indexOrBeforeSheet) {
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

        return this;
    }

    /**
     * Generates the workbook output.
     * @param {string} [type] - The type of the data to return: base64, binarystring, uint8array, arraybuffer, blob, nodebuffer. Defaults to 'nodebuffer' in Node.js and 'blob' in browsers.
     * @returns {string|Uint8Array|ArrayBuffer|Blob|Buffer} The data.
     *//**
     * Generates the workbook output.
     * @param {{}} [opts] Options
     * @param {string} [opts.type] - The type of the data to return: base64, binarystring, uint8array, arraybuffer, blob, nodebuffer. Defaults to 'nodebuffer' in Node.js and 'blob' in browsers.
     * @param {string} [opts.password] - The password to use to encrypt the workbook.
     * @returns {string|Uint8Array|ArrayBuffer|Blob|Buffer} The data.
     */
    outputAsync(opts) {
        opts = opts || {};
        if (typeof opts === 'string') opts = { type: opts };

        this._setSheetRefs();

        this._sheetsNode.children = [];
        this._sheets.forEach((sheet, i) => {
            const sheetPath = `xl/worksheets/sheet${i + 1}.xml`;
            const sheetRelsPath = `xl/worksheets/_rels/sheet${i + 1}.xml.rels`;
            const sheetXmls = sheet.toXmls();
            const relationship = this._relationships.findById(sheetXmls.id.attributes['r:id']);
            relationship.attributes.Target = `worksheets/sheet${i + 1}.xml`;
            this._sheetsNode.children.push(sheetXmls.id);
            this._zip.file(sheetPath, xmlBuilder.build(sheetXmls.sheet), zipFileOpts);

            const relationshipsXml = xmlBuilder.build(sheetXmls.relationships);
            if (relationshipsXml) {
                this._zip.file(sheetRelsPath, relationshipsXml, zipFileOpts);
            } else {
                this._zip.remove(sheetRelsPath);
            }
        });

        // Set the app security to true if a password is set, false if not.
        // this._appProperties.isSecure(!!opts.password);

        // Convert the various components to XML strings and add them to the zip.
        this._zip.file("[Content_Types].xml", xmlBuilder.build(this._contentTypes), zipFileOpts);
        this._zip.file("docProps/app.xml", xmlBuilder.build(this._appProperties), zipFileOpts);
        this._zip.file("xl/_rels/workbook.xml.rels", xmlBuilder.build(this._relationships), zipFileOpts);
        this._zip.file("xl/sharedStrings.xml", xmlBuilder.build(this._sharedStrings), zipFileOpts);
        this._zip.file("xl/styles.xml", xmlBuilder.build(this._styleSheet), zipFileOpts);
        this._zip.file("xl/workbook.xml", xmlBuilder.build(this._node), zipFileOpts);

        // Generate the zip.
        return this._zip.generateAsync({
            type: "nodebuffer",
            compression: "DEFLATE"
        }).then(output => {
            // If a password is set, encrypt the workbook.
            if (opts.password) output = encryptor.encrypt(output, opts.password);

            // Convert and return
            return this._convertBufferToOutput(output, opts.type);
        });
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
        return this._sheets.slice();
    }

    /**
     * Write the workbook to file. (Not supported in browsers.)
     * @param {string} path - The path of the file to write.
     * @param {{}} [opts] - Options
     * @param {string} [opts.password] - The password to encrypt the workbook.
     * @returns {Promise.<undefined>} A promise.
     */
    toFileAsync(path, opts) {
        if (process.browser) throw new Error("Workbook.toFileAsync is not supported in the browser.");
        return this.outputAsync(opts)
            .then(data => new externals.Promise((resolve, reject) => {
                fs.writeFile(path, data, err => {
                    if (err) return reject(err);
                    resolve();
                });
            }));
    }

    /**
     * Gets a scoped defined name.
     * @param {Sheet} sheetScope - The sheet the name is scoped to. Use undefined for workbook scope.
     * @param {string} name - The defined name.
     * @returns {undefined|Cell|Range|Row|Column} What the defined name refers to.
     * @ignore
     *//**
     * Sets a scoped defined name.
     * @param {Sheet} sheetScope - The sheet the name is scoped to. Use undefined for workbook scope.
     * @param {string} name - The defined name.
     * @param {undefined|Cell|Range|Row|Column} refersTo - What the defined name refers to.
     * @returns {Workbook} The workbook.
     * @ignore
     */
    scopedDefinedName(sheetScope, name, refersTo) {
        let definedNamesNode = xmlq.findChild(this._node, "definedNames");
        let definedNameNode = definedNamesNode && _.find(definedNamesNode.children, node => node.attributes.name === name && node.localSheet === sheetScope);

        return new ArgHandler('Workbook.scopedDefinedName')
            .case(['*', 'string'], () => {
                // Get the address from the definedNames node.
                const refersTo = definedNameNode && definedNameNode.children[0];
                if (!refersTo) return undefined;

                // Try to parse the address.
                const ref = addressConverter.fromAddress(refersTo);
                if (!ref) return refersTo;

                // Load the appropriate selection type.
                const sheet = this.sheet(ref.sheetName);
                if (ref.type === 'cell') return sheet.cell(ref.rowNumber, ref.columnNumber);
                if (ref.type === 'range') return sheet.range(ref.startRowNumber, ref.startColumnNumber, ref.endRowNumber, ref.endColumnNumber);
                if (ref.type === 'row') return sheet.row(ref.rowNumber);
                if (ref.type === 'column') return sheet.column(ref.columnNumber);
                return refersTo;
            })
            .case(['*', 'string', 'nil'], () => {
                if (definedNameNode) xmlq.removeChild(definedNamesNode, definedNameNode);
                if (definedNamesNode && !definedNamesNode.children.length) xmlq.removeChild(this._node, definedNamesNode);
                return this;
            })
            .case(['*', 'string', '*'], () => {
                if (typeof refersTo !== 'string') {
                    refersTo = refersTo.address({
                        includeSheetName: true,
                        anchored: true
                    });
                }

                if (!definedNamesNode) {
                    definedNamesNode = {
                        name: "definedNames",
                        attributes: {},
                        children: []
                    };

                    xmlq.insertInOrder(this._node, definedNamesNode, nodeOrder);
                }

                if (!definedNameNode) {
                    definedNameNode = {
                        name: "definedName",
                        attributes: { name },
                        children: [refersTo]
                    };

                    if (sheetScope) definedNameNode.localSheet = sheetScope;

                    xmlq.appendChild(definedNamesNode, definedNameNode);
                }
                
                definedNameNode.children = [refersTo];
                
                return this;
            })
            .handle(arguments);
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
     * @param {string|ArrayBuffer|Uint8Array|Buffer|Blob} data - The data to load.
     * @param {{}} [opts] - Options
     * @param {boolean} [opts.base64=false] - No used unless input is a string. True if the input string is base64 encoded, false for binary.
     * @returns {Promise.<Workbook>} The workbook.
     * @private
     */
    _initAsync(data, opts) {
        opts = opts || {};

        this._maxSheetId = 0;
        this._sheets = [];

        return externals.Promise.resolve()
            .then(() => {
                // Make sure the input is a Buffer
                return this._convertInputToBufferAsync(data, opts.base64)
                    .then(buffer => {
                        data = buffer;
                    });
            })
            .then(() => {
                if (!opts.password) return;
                return encryptor.decryptAsync(data, opts.password)
                    .then(decrypted => {
                        data = decrypted;
                    });
            })
            .then(() => JSZip.loadAsync(data))
            .then(zip => {
                this._zip = zip;
                return this._parseNodesAsync([
                    "[Content_Types].xml",
                    "docProps/app.xml",
                    "xl/_rels/workbook.xml.rels",
                    "xl/sharedStrings.xml",
                    "xl/styles.xml",
                    "xl/workbook.xml"
                ]);
            })
            .then(nodes => {
                const contentTypesNode = nodes[0];
                const appPropertiesNode = nodes[1];
                const relationshipsNode = nodes[2];
                const sharedStringsNode = nodes[3];
                const styleSheetNode = nodes[4];
                const workbookNode = nodes[5];

                // Load the various components.
                this._contentTypes = new ContentTypes(contentTypesNode);
                this._appProperties = new AppProperties(appPropertiesNode);
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
            .then(() => this._parseSheetRefs())
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

    /**
     * Parse the sheet references out so we can reorder freely.
     * @returns {undefined}
     * @private
     */
    _parseSheetRefs() {
        // Parse the active sheet.
        const bookViewsNode = xmlq.findChild(this._node, "bookViews");
        const workbookViewNode = bookViewsNode && xmlq.findChild(bookViewsNode, "workbookView");
        const activeTabId = workbookViewNode && workbookViewNode.attributes.activeTab || 0;
        this._activeSheet = this._sheets[activeTabId];

        // Set the location sheet on the defined name nodes. The defined name should point to the index of the sheet
        // but reordering sheets messes this up. So store it on the node and we'll update the index on XML build.
        const definedNamesNode = xmlq.findChild(this._node, "definedNames");
        if (definedNamesNode) {
            _.forEach(definedNamesNode.children, definedNameNode => {
                if (definedNameNode.attributes.hasOwnProperty("localSheetId")) {
                    definedNameNode.localSheet = this._sheets[definedNameNode.attributes.localSheetId];
                }
            });
        }
    }

    /**
     * Set the proper sheet references in the XML.
     * @returns {undefined}
     * @private
     */
    _setSheetRefs() {
        // Set the active sheet.
        let bookViewsNode = xmlq.findChild(this._node, "bookViews");
        if (!bookViewsNode) {
            bookViewsNode = { name: 'bookViews', attributes: {}, children: [] };
            xmlq.insertInOrder(this._node, bookViewsNode, nodeOrder);
        }

        let workbookViewNode = xmlq.findChild(bookViewsNode, "workbookView");
        if (!workbookViewNode) {
            workbookViewNode = { name: 'workbookView', attributes: {}, children: [] };
            xmlq.appendChild(bookViewsNode, workbookViewNode);
        }

        workbookViewNode.attributes.activeTab = this._sheets.indexOf(this._activeSheet);

        // Set the defined names local sheet indexes.
        const definedNamesNode = xmlq.findChild(this._node, "definedNames");
        if (definedNamesNode) {
            _.forEach(definedNamesNode.children, definedNameNode => {
                if (definedNameNode.localSheet) {
                    definedNameNode.attributes.localSheetId = this._sheets.indexOf(definedNameNode.localSheet);
                }
            });
        }
    }

    /**
     * Convert buffer to desired output format
     * @param {Buffer} buffer - The buffer
     * @param {string} type - The type to convert to: buffer/nodebuffer, blob, base64, binarystring, uint8array, arraybuffer
     * @returns {Buffer|Blob|string|Uint8Array|ArrayBuffer} The output
     * @private
     */
    _convertBufferToOutput(buffer, type) {
        if (!type) type = process.browser ? "blob" : "nodebuffer";

        if (type === "buffer" || type === "nodebuffer") return buffer;
        if (process.browser && type === "blob") return new Blob([buffer], { type: Workbook.MIME_TYPE });
        if (type === "base64") return buffer.toString("base64");
        if (type === "binarystring") return buffer.toString("utf8");
        if (type === "uint8array") return new Uint8Array(buffer);
        if (type === "arraybuffer") return new Uint8Array(buffer).buffer;

        throw new Error(`Output type '${type}' not supported.`);
    }

    /**
     * Convert input to buffer
     * @param {Buffer|Blob|string|Uint8Array|ArrayBuffer} input - The input
     * @param {boolean} [base64=false] - Only applies if input is a string. If true, the string is base64 encoded, false for binary
     * @returns {Promise.<Buffer>} The buffer.
     * @private
     */
    _convertInputToBufferAsync(input, base64) {
        return externals.Promise.resolve()
            .then(() => {
                if (Buffer.isBuffer(input)) return input;

                if (process.browser && input instanceof Blob) {
                    return new externals.Promise(resolve => {
                        const fileReader = new FileReader();
                        fileReader.onload = event => {
                            resolve(Buffer.from(event.target.result));
                        };
                        fileReader.readAsArrayBuffer(input);
                    });
                }

                if (typeof input === "string" && base64) return Buffer.from(input, "base64");
                if (typeof input === "string" && !base64) return Buffer.from(input, "utf8");
                if (input instanceof Uint8Array || input instanceof ArrayBuffer) return Buffer.from(input);

                throw new Error(`Input type unknown.`);
            });
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
