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

class Drawing {

    constructor(workbook, node, relationshipNode) {
        this._node = node;
        this._relationship = relationshipNode;
        this._init(workbook, node, relationshipNode)
    }

    _init(workbook, node, relationshipNode) {

        console.log(node)
        console.log(relationshipNode)



        //basic
        this.id = relationshipNode.attributes.Id;
        this.name;
        this.description;
        this.path = relationshipNode.attributes.Target;

        //position
        this.fromCol;
        this.toCol;
        this.fromRow;
        this.toRow;

        //offset
        this.fromColOff;
        this.toColOff;
        this.fromRowOff;
        this.toRowOff;
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
            .case('*', path => {
                tmpzip = new JSZip();
                tmpzip.file('image', fs.readFileSync(path, {encoding: 'binary'})).generateAsync({type : "string"}).then(() => {
                    this.workbook._zip.files[this.path.replace('..', 'xl')]._data = tmpzip.files.image._data;
                }).then(() => {
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
    * @returns {{}} The XML forms.
    * @ignore
    */
    toXmls() {
        // Shallow clone the node so we don't have to remove these children later if they don't belong.
        const node = _.clone(this._node);
        node.children = node.children.slice();

        // Add the columns if needed.
        this._colsNode.children = _.filter(this._colNodes, (colNode, i) => {
            // Columns should only be present if they have attributes other than min/max.
            return colNode && i === colNode.attributes.min && Object.keys(colNode.attributes).length > 2;
        });
        if (this._colsNode.children.length) {
            xmlq.insertInOrder(node, this._colsNode, nodeOrder);
        }

        // Add the hyperlinks if needed.
        this._hyperlinksNode.children = _.values(this._hyperlinks);
        if (this._hyperlinksNode.children.length) {
            xmlq.insertInOrder(node, this._hyperlinksNode, nodeOrder);
        }

        // Add the merge cells if needed.
        this._mergeCellsNode.children = _.values(this._mergeCells);
        if (this._mergeCellsNode.children.length) {
            xmlq.insertInOrder(node, this._mergeCellsNode, nodeOrder);
        }

        // Add the DataValidation cells if needed.
        this._dataValidationsNode.children = _.values(this._dataValidations);
        if (this._dataValidationsNode.children.length) {
            xmlq.insertInOrder(node, this._dataValidationsNode, nodeOrder);
        }

        return {
            id: this._idNode,
            sheet: node,
            relationships: this._relationships
        };
    }

}


module.exports = Drawing;