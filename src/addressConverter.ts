const ADDRESS_REGEX = /^(?:'?(.+?)'?!)?(?:(\$)?([A-Z]+)(\$)?(\d+)(?::(\$)?([A-Z]+)(\$)?(\d+))?|(\$)?([A-Z]+):(\$)?([A-Z]+)|(\$)?(\d+):(\$)?(\d+))$/;

/**
 * Convert a column name to a number.
 * @param {string} name - The column name.
 * @returns {number} The number.
 */
export function columnNameToNumber(name: string): number {
    name = name.toUpperCase();
    let sum = 0;
    for (let i = 0; i < name.length; i++) {
        sum = sum * 26;
        sum = sum + (name[i].charCodeAt(0) - 'A'.charCodeAt(0) + 1);
    }

    return sum;
}

/**
 * Convert a column number to a name.
 * @param {number} number - The column number.
 * @returns {string} The name.
 */
export function columnNumberToName(number: number): string {
    let dividend = number;
    let name = '';
    let modulo = 0;

    while (dividend > 0) {
        modulo = (dividend - 1) % 26;
        name = String.fromCharCode('A'.charCodeAt(0) + modulo) + name;
        dividend = Math.floor((dividend - modulo) / 26);
    }

    return name;
}

/**
 * Convert an address to a reference object.
 * @param {string} address - The address.
 * @returns {{}} The reference object.
 */
export function fromAddress(address: string): any { // TODO
    const match = address.match(ADDRESS_REGEX);
    if (!match) return;
    // if (!match) throw new Error(`Address "${address}" is not valid.`);
    let ref: Address | undefined;

    if (match[3] && match[7]) {
        const startColumnName = match[3];
        const endColumnName = match[7]
        ref = {
            type: 'range',
            startColumnAnchored: !!match[2],
            startColumnName,
            startColumnNumber: columnNameToNumber(startColumnName),
            startRowAnchored: !!match[4],
            startRowNumber: parseInt(match[5]),
            endColumnAnchored: !!match[6],
            endColumnName,
            endColumnNumber: columnNameToNumber(endColumnName),
            endRowAnchored: !!match[8],
            endRowNumber: parseInt(match[9]),
        };
    } else if (match[3]) {
        const columnName = match[3];
        ref = {
            type: 'cell',
            columnAnchored: !!match[2],
            columnName,
            columnNumber: columnNameToNumber(columnName),
            rowAnchored: !!match[4],
            rowNumber: parseInt(match[5]),
        };
    } else if (match[11] && match[11] !== match[13]) {
        const startColumnName = match[11];
        const endColumnName = match[13];
        ref = {
            type: 'columnRange',
            startColumnAnchored: !!match[10],
            startColumnName,
            startColumnNumber: columnNameToNumber(startColumnName),
            endColumnAnchored: !!match[12],
            endColumnName,
            endColumnNumber: columnNameToNumber(endColumnName),
        };
    } else if (match[11]) {
        const columnName = match[11];
        ref = {
            type: 'column',
            columnAnchored: !!match[10],
            columnName,
            columnNumber: columnNameToNumber(columnName),
        };
    } else if (match[15] && match[15] !== match[17]) {
        ref = {
            type: 'rowRange',
            startRowAnchored: !!match[14],
            startRowNumber: parseInt(match[15]),
            endRowAnchored: !!match[16],
            endRowNumber: parseInt(match[17]),
        };
    } else if (match[15]) {
        ref = {
            type: 'row',
            rowAnchored: !!match[14],
            rowNumber: parseInt(match[15]),
        }
    }

    if (!ref) throw new Error("Unsupported address type.");

    if (match[1]) ref.sheetName = match[1].replace(/''/g, "'");

    return ref;
}

interface CellAddress {
    type: 'cell';
    sheetName?: string;
    columnName: string;
    columnNumber: number;
    columnAnchored: boolean;
    rowNumber: number;
    rowAnchored: boolean;
}

interface RangeAddress {
    type: 'range';
    sheetName?: string;
    startColumnAnchored: boolean;
    startColumnName: string;
    startColumnNumber: number;
    startRowAnchored: boolean;
    startRowNumber: number;
    endColumnAnchored: boolean;
    endColumnName: string;
    endColumnNumber: number;
    endRowAnchored: boolean;
    endRowNumber: number;
}

interface ColumnRangeAddress {
    type: 'columnRange';
    sheetName?: string;
    startColumnAnchored: boolean;
    startColumnName: string;
    startColumnNumber: number;
    endColumnAnchored: boolean;
    endColumnName: string;
    endColumnNumber: number;
}

interface ColumnAddress {
    type: 'column';
    sheetName?: string;
    columnAnchored: boolean;
    columnName: string;
    columnNumber: number;
}

interface RowRangeAddress {
    type: 'rowRange';
    sheetName?: string;
    startRowAnchored: boolean;
    startRowNumber: number;
    endRowAnchored: boolean;
    endRowNumber: number;
}

interface RowAddress {
    type: 'row';
    sheetName?: string;
    rowAnchored: boolean;
    rowNumber: number;
}

type Address = CellAddress | RangeAddress | ColumnRangeAddress | ColumnAddress | RowRangeAddress | RowAddress;

/**
 * Convert a reference object to an address.
 * @param {{}} ref - The reference object.
 * @returns {string} The address.
 */
export function toAddress(ref: Address) {
    let a: any, b: any;
    const sheetName = ref.sheetName;

    if (ref.type === 'cell') {
        a = {
            columnName: ref.columnName,
            columnNumber: ref.columnNumber,
            columnAnchored: ref.columnAnchored,
            rowNumber: ref.rowNumber,
            rowAnchored: ref.rowAnchored
        };
    } else if (ref.type === 'range') {
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
        a = b = {
            columnName: ref.columnName,
            columnNumber: ref.columnNumber,
            columnAnchored: ref.columnAnchored
        };
    } else if (ref.type === 'row') {
        a = b = {
            rowNumber: ref.rowNumber,
            rowAnchored: ref.rowAnchored
        };
    } else if (ref.type === 'columnRange') {
        a = {
            columnName: ref.startColumnName,
            columnNumber: ref.startColumnNumber,
            columnAnchored: ref.startColumnAnchored
        };
        b = {
            columnName: ref.endColumnName,
            columnNumber: ref.endColumnNumber,
            columnAnchored: ref.endColumnAnchored
        };
    } else if (ref.type === 'rowRange') {
        a = {
            rowNumber: ref.startRowNumber,
            rowAnchored: ref.startRowAnchored
        };
        b = {
            rowNumber: ref.endRowNumber,
            rowAnchored: ref.endRowAnchored
        };
    }

    let address = '';
    if (sheetName) address = `${address}'${sheetName.replace(/'/g, "''")}'!`;
    if (a.columnAnchored) address = `${address}$`;
    if (a.columnName) address = address + a.columnName;
    else if (a.columnNumber) address = address + columnNumberToName(a.columnNumber);
    if (a.rowAnchored) address = `${address}$`;
    if (a.rowNumber) address = address + a.rowNumber;

    if (b) {
        address = `${address}:`;
        if (b.columnAnchored) address = `${address}$`;
        if (b.columnName) address = address + b.columnName;
        else if (b.columnNumber) address = address + columnNumberToName(b.columnNumber);
        if (b.rowAnchored) address = `${address}$`;
        if (b.rowNumber) address = address + b.rowNumber;
    }

    return address;
}

