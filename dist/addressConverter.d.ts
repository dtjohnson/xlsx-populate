/**
 * Convert a column name to a number.
 * @param {string} name - The column name.
 * @returns {number} The number.
 */
export declare function columnNameToNumber(name: string): number;
/**
 * Convert a column number to a name.
 * @param {number} number - The column number.
 * @returns {string} The name.
 */
export declare function columnNumberToName(number: number): string;
/**
 * Convert an address to a reference object.
 * @param {string} address - The address.
 * @returns {{}} The reference object.
 */
export declare function fromAddress(address: string): any;
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
declare type Address = CellAddress | RangeAddress | ColumnRangeAddress | ColumnAddress | RowRangeAddress | RowAddress;
/**
 * Convert a reference object to an address.
 * @param {{}} ref - The reference object.
 * @returns {string} The address.
 */
export declare function toAddress(ref: Address): string;
export {};
//# sourceMappingURL=addressConverter.d.ts.map