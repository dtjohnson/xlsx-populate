"use strict";

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

    toAddress(ref) {
        let a, b;

        if (ref.type === 'column') {
            a = b = {
                columnName: ref.columnName,
                columnAnchored: ref.anchored,
                sheetName: ref.sheetName
            };
        }

        let address = '';
        if (a.sheetName) address += `${a.sheetName}!`;
        if (a.columnAnchored) address += '$';
        if (a.columnName) address += a.columnName;
        if (a.rowAnchored) address += '$';
        if (a.rowNumber) address += a.rowNumber;

        if (b) {
            address += ':';
            if (b.columnAnchored) address += '$';
            if (b.columnName) address += b.columnName;
            if (b.rowAnchored) address += '$';
            if (b.rowNumber) address += b.rowNumber;
        }

        return address;
    },

    fromAddress(address) {

    }
};
