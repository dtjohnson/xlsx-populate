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

    constructor(workbook, node, relationshipNode) {
        this._node = node;
        this._relationshipNode = relationshipNode;
        this._init(workbook, node, relationshipNode)
    }

    _init(workbook, node, relationshipNode) {

        let to = node.children[0].children;
        let from = node.children[1].children;
        let pic = node.children[2].children[0].children[0];

        //basic
        this.id = relationshipNode.attributes.Id;
        this.name = pic.attributes.name;
        this.description = pic.attributes.descr ? pic.attributes.descr : '';
        this.title = pic.attributes.title ? pic.attributes.title : '';
        this.path = relationshipNode.attributes.Target;

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

        //clear the memory
        to=null;
        from=null;
        pic=null;
    }


    id() {
        return new ArgHandler("Drawing.id")
            .case(() => {
                return this.id;
            })
        .handle(arguments);
    }

    name() {
        return new ArgHandler("Drawing.name")
            .case('string', name => {
                this.name = name;
                return this;
            })
            .case(() => {
                return this.name;
            })
        .handle(arguments);
    }

    image() {
        return new ArgHandler("Drawing.image")
            .case('string', path => {
                tmpzip = new JSZip();
                tmpzip.file('image', fs.readFileSync(path, {encoding: 'binary'})).generateAsync({type : "string"}).then(() => {
                    this.workbook._zip.files[this.path.replace('..', 'xl')]._data = tmpzip.files.image._data;
                }).then(() => {
                    unset(tmpzip);
                    return this;
                });
            })
            .case(() => {
                this.workbook._zip.files[this.path.replace('..', 'xl')]
            })
        .handle(arguments);
    }

    description() {
        return new ArgHandler("Drawing.description")
            .case('string', description => {
                this.description = description;
                return this;
            })
            .case(() => {
                return this.description;
            })
        .handle(arguments);
    }

    path() {
        return new ArgHandler("Drawing.path")
            .case('string', path => {
                this.path = path;
                return this;
            })
            .case(() => {
                return this.path;
            })
        .handle(arguments);
    }

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

    /* FROM Sheet.js  */

    /**
    * Convert the sheet to a collection of XML objects.
    * @returns {{}} The XML nodes.
    * @ignore
    */
    toXml() {
        this._relationshipNode.attributes.Id = this.id;
        this._relationshipNode.attributes.Target = this.path;

        this._node.children[0].children[0].children[0] = this.toCol;
        this._node.children[0].children[1].children[0] = this.toColOff;
        this._node.children[0].children[2].children[0] = this.toRow;
        this._node.children[0].children[3].children[0] = this.toRowOff;

        this._node.children[1].children[0].children[0] = this.fromCol;
        this._node.children[1].children[1].children[0] = this.fromColOff;
        this._node.children[1].children[2].children[0] = this.fromRow;
        this._node.children[1].children[3].children[0] = this.fromRowOff;

        this._node.children[2].children[0].children[0].attributes.name = this.name;
        this._node.children[2].children[0].children[0].attributes.descr = this.description;
        this._node.children[2].children[0].children[0].attributes.title = this.title;

        return [ this._node, this._relationshipNode ];
    }

}


module.exports = Drawing;