"use strict";

const _ = require("lodash");
const ADDRESS_REGEX = /^(?:'?(.+?)'?!)?(?:(\$)?([A-Z]+)(\$)?(\d+)(?::(\$)?([A-Z]+)(\$)?(\d+))?|(\$)?([A-Z]+):(\$)?([A-Z]+)|(\$)?(\d+):(\$)?(\d+))$/;

module.exports = {
    columnNumberToName(number) {
        let dividend = number;
        let name = '';
        let modulo = 0;

        while (dividend > 0) {
            modulo = (dividend - 1) % 26;
            name = String.fromCharCode('A'.charCodeAt(0) + modulo) + name;
            dividend = Math.floor((dividend - modulo) / 26);
        }

        return name;
    },

    columnNameToNumber(name) {
        if (!name || typeof name !== "string") return;

        name = name.toUpperCase();
        let sum = 0;
        for (let i = 0; i < name.length; i++) {
            sum *= 26;
            sum += name[i].charCodeAt(0) - 'A'.charCodeAt(0) + 1;
        }

        return sum;
    },

    toAddress(ref) {
        let sheetName, a, b;

        if (ref.type === 'cell') {
            sheetName = ref.sheetName;
            a = {
                columnName: ref.columnName,
                columnNumber: ref.columnNumber,
                columnAnchored: ref.columnAnchored,
                rowNumber: ref.rowNumber,
                rowAnchored: ref.rowAnchored
            };
        } else if (ref.type === 'range') {
            sheetName = ref.sheetName;
            a = {
                columnName: ref.startColumnName,
                columnNumber: ref.startColumnNumber,
                columnAnchored: ref.startColumnAnchored,
                rowNumber: ref.startRowNumber,
                rowAnchored: ref.startRowAnchored
            };
            b = {
                columnName: ref.endColumnName,
                columnNumber: ref.endColumnNumber,
                columnAnchored: ref.endColumnAnchored,
                rowNumber: ref.endRowNumber,
                rowAnchored: ref.endRowAnchored
            };
        } else if (ref.type === 'column') {
            sheetName = ref.sheetName;
            a = b = {
                columnName: ref.columnName,
                columnNumber: ref.columnNumber,
                columnAnchored: ref.anchored
            };
        } else if (ref.type === 'row') {
            sheetName = ref.sheetName;
            a = b = {
                rowNumber: ref.rowNumber,
                rowAnchored: ref.anchored
            };
        }

        let address = '';
        if (sheetName) address += `${sheetName}!`;
        if (a.columnAnchored) address += '$';
        if (a.columnName) address += a.columnName;
        else if (a.columnNumber) address += this.columnNumberToName(a.columnNumber);
        if (a.rowAnchored) address += '$';
        if (a.rowNumber) address += a.rowNumber;

        if (b) {
            address += ':';
            if (b.columnAnchored) address += '$';
            if (b.columnName) address += b.columnName;
            else if (b.columnNumber) address += this.columnNumberToName(b.columnNumber);
            if (b.rowAnchored) address += '$';
            if (b.rowNumber) address += b.rowNumber;
        }

        return address;
    },

    fromAddress(address) {
        const parts = address.split(',');
        if (parts.length > 1) return {
            type: 'group',
            refs: parts.map(part => this.fromAddress(part))
        };

        const match = address.match(ADDRESS_REGEX);
        if (!match) throw new Error(`Unable to parse address "${address}"`);
        const ref = {};

        if (match[1]) ref.sheetName = match[1];
        if (match[3] && match[7]) {
            ref.type = 'range';
            ref.startColumnAnchored = !!match[2];
            ref.startColumnName = match[3];
            ref.startColumnNumber = this.columnNameToNumber(ref.startColumnName);
            ref.startRowAnchored = !!match[4];
            ref.startRowNumber = parseInt(match[5]);
            ref.endColumnAnchored = !!match[6];
            ref.endColumnName = match[7];
            ref.endColumnNumber = this.columnNameToNumber(ref.endColumnName);
            ref.endRowAnchored = !!match[8];
            ref.endRowNumber = parseInt(match[9]);
        } else if (match[3]) {
            ref.type = 'cell';
            ref.columnAnchored = !!match[2];
            ref.columnName = match[3];
            ref.columnNumber = this.columnNameToNumber(ref.columnName);
            ref.rowAnchored = !!match[4];
            ref.rowNumber = parseInt(match[5]);
        } else if (match[11] && match[11] === match[13]) {
            ref.type = 'columnRange';
            ref.startColumnAnchored = !!match[10];
            ref.startColumnName = match[11];
            ref.startColumnNumber = this.columnNameToNumber(ref.startColumnName);
            ref.endColumnAnchored = !!match[12];
            ref.endColumnName = match[13];
            ref.endColumnNumber = this.columnNameToNumber(ref.endColumnName);
        } else if (match[11]) {
            ref.type = 'column';
            ref.columnAnchored = !!match[10];
            ref.columnName = match[11];
            ref.columnNumber = this.columnNameToNumber(ref.columnName);
        } else if (match[15] && match[15] === match[17]) {
            ref.type = 'rowRange';
            ref.startRowAnchored = !!match[14];
            ref.startRowNumber = parseInt(match[15]);
            ref.endRowAnchored = !!match[16];
            ref.endRowNumber = parseInt(match[17]);
        } else if (match[15]) {
            ref.type = 'row';
            ref.rowAnchored = !!match[14];
            ref.rowNumber = parseInt(match[15]);
        }

        return ref;
    }
};
