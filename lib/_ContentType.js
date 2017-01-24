"use strict";

// TODO: Tests
// TODO: JSDoc
// TODO: Debugs

class _ContentType {
    constructor(node) {
        this._node = node;
    }

    partName() {
        return this._node.getAttribute("PartName");
    }

    contentType() {
        return this._node.getAttribute("ContentType");
    }
}

module.exports = _ContentType;

/*
 <Default Extension="xml" ContentType="application/xml"/>

 <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
 */
