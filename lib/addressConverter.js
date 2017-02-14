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

    }
};
