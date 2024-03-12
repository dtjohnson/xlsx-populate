"use strict";

const ArgHandler = require("./ArgHandler");
const addressConverter = require("./addressConverter");
const JSZip = require('jszip');

class Drawing {
    constructor(nodeId, workbook, sheet, node, relationshipNode) {
        this._nodeId = nodeId;
        this._workbook = workbook;
        this._sheet = sheet;
        this._node = node;
        this._relationship = relationshipNode;
        this._init(node);
    }

    _init(node) {
        let to = node.children[1].children;
        let from = node.children[0].children;
        let pic = node.children[2].children[0].children[0];

        if (this._relationship) {
            if (this._relationship.attributes.Type.toString().includes('/image')) {
                this.returnImage = true;

                // Basic
                this._id = node.children[2].children[1].children[0].attributes['r:embed'];
                this._name = pic.attributes.name;
                this._description = pic.attributes.descr ? pic.attributes.descr : '';
                this._title = pic.attributes.title ? pic.attributes.title : '';
                this._path = this._relationship.attributes.Target;
                this.compressedSize = this._workbook._zip.files[this._path.replace('..', 'xl')]._data.compressedSize;
                this.uncompressedSize = this._workbook._zip.files[this._path.replace('..', 'xl')]._data.uncompressedSize;

                // Position
                this.fromCol = from[0].children[0];
                this.toCol = to[0].children[0];
                this.fromRow = from[2].children[0];
                this.toRow = to[2].children[0];

                // Offset
                this.fromColOff = from[1].children[0];
                this.toColOff = to[1].children[0];
                this.fromRowOff = from[3].children[0];
                this.toRowOff = to[3].children[0];
            } else {
                this.returnImage = true;
            }
        } else {
            this.returnImage = true;
        }

        // Clear the memory
        to = null;
        from = null;
        pic = null;
    }

    /**
     * Get Drawing ID
     * @returns {String} - Drawing ID
     * @ignore
     */
    id() {
        return new ArgHandler("Drawing.id")
            .case(() => {
                return this._id;
            })
            .handle(arguments);
    }

    /**
     * Set the Drawing Defined Name
     * @param newName - the new name
     * @returns {boolean} - true
     */
    /**
    * Get the Drawing Defined Name
    * @returns {String} - Drawings Name
    */
    name() {
        return new ArgHandler("Drawing.name")
            .case(() => {
                return this._name;
            })
            .case('string', newName => {
                this._name = newName;
                return true;
            })
            .handle(arguments);
    }

    /**
     * Replace the image with a new image
     * @param imagePath - Path to the new image.
     * @returns {Drawing} - drawing
     */
    /** 
    * Get the image Data from the zip
    * @returns {Object} - image data
    */
    image() {
        return new ArgHandler("Drawing.image")
            .case(() => {
                return this._workbook._zip.files[this._path.replace('..', 'xl')];
            })
            .case('blob', file => {
                this._workbook._zip.files[this._path.replace('..', 'xl')]._data = new JSZip().file('i', file).files.i._data;
                return this;
            })
            .handle(arguments);
    }

    /**
     * Set the Description for the Drawing
     * @param description - the new description
     * @returns {Drawing} - drawing
     *//** 
    * Get the Description for the Drawing
    * @returns {String} - description
    */
    description() {
        return new ArgHandler("Drawing.description")
            .case(() => {
                return this._description;
            })
            .case('string', desc => {
                this._description = desc;
                return this;
            })
            .handle(arguments);
    }

    /**
     * Set the Title for the Drawing
     * @param newTitle - the new title for the Drawing
     * @returns {Drawing} - drawing
     *//** 
    * Get the Title for the Drawing
    * @returns {String} - title
    */
    title() {
        return new ArgHandler("Drawing.title")
            .case(() => {
                return this._title;
            })
            .case('string', newTitle => {
                this._title = newTitle;
                return this;
            })
            .handle(arguments);
    }

    /**
     * Set the Image Path
     * @param path - new Path inside the xlsx file
     * @returns {Drawing} - drawing
     *//** 
    * Get the Image Path
    * @returns {String} - path
    */
    path() {
        return new ArgHandler("Drawing.path")
            .case(() => {
                return this._path;
            })
            .case('string', path => {
                this._path = path;
                return this;
            })
            .handle(arguments);
    }

    /**
     * get the From data from the Drawing
     */
    from() {
        return new ArgHandler("Drawing.from")
            .case(() => {
                return {
                    col: this.fromCol,
                    row: this.fromRow,
                    colOffset: this.fromColOff,
                    rowOffset: this.fromRowOff
                };
            })
            .case('string', address => {
                const ref = addressConverter.fromAddress(address);
                if (ref.type !== 'cell') throw new Error('Sheet.cell: Invalid address.');
                this.fromCol = ref.columnNumber - 1;
                this.fromRow = ref.rowNumber - 1;
                this.fromColOff = 0;
                this.fromRowOff = 0;
                return this;
            })
            .handle(arguments);
    }

    /**
     * Get the To Data from the Drawing
     */
    to() {
        return new ArgHandler("Drawing.to")
            .case(() => {
                return {
                    col: this.toCol,
                    row: this.toRow,
                    colOffset: this.toColOff,
                    rowOffset: this.toRowOff
                };
            })
            .case('string', address => {
                const ref = addressConverter.fromAddress(address);
                if (ref.type !== 'cell') throw new Error('Sheet.cell: Invalid address.');
                this.toCol = ref.columnNumber;
                this.toRow = ref.rowNumber;
                this.toColOff = 0;
                this.toRowOff = 0;
                return this;
            })
            .handle(arguments);
    }

    /**
     * Get the image file size
     */
    size() {
        return new ArgHandler("Drawing.size")
            .case(() => {
                return this.uncompressedSize;
            })
            .handle(arguments);
    }

    /**
    * overrides the original values in the XML, and then return the XML objects.
    * @returns {{}} The XML nodes.
    * @ignore
    */
    toXml() {
        if (this.returnImage) {
            this._relationship.attributes.Id = this._id;
            this._relationship.attributes.Target = this._path;

            this._node.children[1].children[0].children[0] = this.toCol;
            this._node.children[1].children[1].children[0] = this.toColOff;
            this._node.children[1].children[2].children[0] = this.toRow;
            this._node.children[1].children[3].children[0] = this.toRowOff;

            this._node.children[0].children[0].children[0] = this.fromCol;
            this._node.children[0].children[1].children[0] = this.fromColOff;
            this._node.children[0].children[2].children[0] = this.fromRow;
            this._node.children[0].children[3].children[0] = this.fromRowOff;

            this._node.children[2].children[0].children[0].attributes.name = this._name;
            this._node.children[2].children[0].children[0].attributes.descr = this._description;
            this._node.children[2].children[0].children[0].attributes.title = this._title;
        }
        return [this._node, this._relationship];
    }
}

module.exports = Drawing;
