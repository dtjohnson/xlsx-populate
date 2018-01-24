"use strict";

const _ = require("lodash");
const Workbook = require("./Workbook");
const Sheet = require("./Sheet");
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
const JSZip = require('jszip');
const fs = require('fs');

class Drawing {

    constructor(workbook, sheet, node, relationshipNode) {
        this._workbook = workbook;
        this._sheet = sheet;
        this._node = node;
        this._relationship = relationshipNode;
        this._init(node, relationshipNode)
    }


    _init(node, relationshipNode) {
        let to = node.children[0].children;
        let from = node.children[1].children;
        let pic = node.children[2].children[0].children[0];
        
        if(this._relationship) {
            if(this._relationship.attributes.Type.toString().includes('/image')) {
                this.returnImage = true;
                //basi
                this._id = node.children[2].children[1].children[0].attributes['r:embed'];
                this._name = pic.attributes.name;
                this._description = pic.attributes.descr ? pic.attributes.descr : '';
                this._title = pic.attributes.title ? pic.attributes.title : '';
                this._path = this._relationship.attributes.Target;

                //position
                this.fromCol = from[0].children[0];
                this.toCol = to[0].children[0];
                this.fromRow = from[2].children[0];
                this.toRow = to[2].children[0];

                //offset
                this.fromColOff = from[1].children[0];
                this.toColOff = to[1].children[0];
                this.fromRowOff = from[3].children[0];
                this.toRowOff = to[3].children[0];
            } else {
                console.log('xlsxPopulate: Err, invalid Drawing Type: ', this._relationship.attributes.Type);
                this.returnImage = true;
            }
        } else {
            this.returnImage = true;
        }

        //clear the memory
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
     * @returns {boolean}
     *//**
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
     * @returns {Drawing}
     *//** 
     * Get the image Data from the zip
     * @returns {Object}
     */
    image() {
        return new ArgHandler("Drawing.image")
            .case(() => {
                return (this._workbook._zip.files[this._path.replace('..', 'xl')]);
            })
            .case('string', imagePath => {
                let file = fs.readFileSync(imagePath, {encoding: 'binary'});
                this._workbook._zip.files[this._path.replace('..', 'xl')]._data = ((new JSZip()).file('i', file).files.i._data);
                return this;
            })
            .handle(arguments);
    }

    /**
     * Set the Description for the Drawing
     * @param description - the new description
     * @returns {Drawing}
     *//** 
     * Get the Description for the Drawing
     * @returns {String}
     * 
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
     * @returns {Drawing}
     *//** 
     * Get the Title for the Drawing
     * @returns {String}
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
     * @returns {Drawing}
     *//** 
     * Get the Image Path
     * @returns {String}
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
            .handle(arguments);
    }

    /**
    * overrides the original values in the XML, and then return the XML objects.
    * @returns {{}} The XML nodes.
    * @ignore
    */
    toXml() {
        if(this.returnImage) {
            this._relationship.attributes.Id = this._id;
            this._relationship.attributes.Target = this._path;
    
            this._node.children[0].children[0].children[0] = this.toCol;
            this._node.children[0].children[1].children[0] = this.toColOff;
            this._node.children[0].children[2].children[0] = this.toRow;
            this._node.children[0].children[3].children[0] = this.toRowOff;
    
            this._node.children[1].children[0].children[0] = this.fromCol;
            this._node.children[1].children[1].children[0] = this.fromColOff;
            this._node.children[1].children[2].children[0] = this.fromRow;
            this._node.children[1].children[3].children[0] = this.fromRowOff;
    
            this._node.children[2].children[0].children[0].attributes.name = this._name;
            this._node.children[2].children[0].children[0].attributes.descr = this._description;
            this._node.children[2].children[0].children[0].attributes.title = this._title;
        }
        return [this._node, this._relationship];
    }

}


module.exports = Drawing;