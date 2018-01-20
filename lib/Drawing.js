"use strict";

const _ = require("lodash");
const Workbook = require("./Workbook");
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

class Drawing {

    constructor(imageNode, relationshipNode) {
        this._init(imageNode, relationshipNode)
    }

    _init(imageNode, relationshipNode) {
        //basic
        this.id;
        this.name;
        this.description;
        this.path;

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

    getId() { return this.id; }
    getName() { return this.name; }
    getDesciption() { return this.description; }
    getPath() { return this.path; }
    getFrom() { return { col: this.fromCol, row: this.fromRow, colOffset: this.fromColOff, rowOffset: this.fromRowOff }; }
    getTo() { return { col: this.toCol, row: this.toRow, colOffset: this.toColOff, rowOffset: this.toRowOff }; }

    replace(newImage) {
        /*
            Replace image with newImage. (takes in path and compresses with jszip and then replaces _data)
        */
        return this;
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