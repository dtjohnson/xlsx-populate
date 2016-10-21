"use strict";

/**
 * Regex to parse Excel addresses.
 * @private
 */
var addressRegex = /^\s*(?:'?(.+?)'?\!)?\$?([A-Z]+)\$?(\d+)\s*$/i;

/**
 * The base date to use for Excel conversion.
 * @private
 */
var dateBase = new Date(1900, 0, 0);

/**
 * The last day in February 1900.
 * @private
 */
var incorrectLeapDate = new Date(1900, 1, 28);

/**
 * Number of milliseconds in a day.
 * @private
 */
var millisecondsInDay = 1000 * 60 * 60 * 24;

module.exports = {
    /**
      * This callback is part of the binarySearch function.
      * @callback binarySearch~getComparableValue
      * @param {Object} value - Element of sortedArray.
      * @returns {Object} Comparable value translated from given value.
      */
    /**
     * This Object is part of the binarySearch function.
     * @typedef {Object} binarySearch~SearchResult
     * @property {number} index - Number of elements less than the targetValue.
     * @property {boolean} found - Indicates whether targetValue can be found at index of sortedArray.
     */
    /**
     * Helper function to search sorted arrays.
     * This is a ranked query and its operation is O(log(n)).
     * @param {*} targetValue - Comparable value to search for within sortedArray.
     * @param {Array} sortedArray - For any indices i and j, if 0 < i < j < A.length then A[i] < A[j] where A is the sortedArray.
     * @param {binarySearch~getComparableValue} getComparableValue - The callback that extracts values from a sortedArray item.
     * @returns {binarySearch~SearchResult} The result of the search.
     */
    binarySearch: function (targetValue, sortedArray, getComparableValue) {
        getComparableValue = getComparableValue || function (value) {
            return value;
        };
        var getValue = function (index) {
            var item = sortedArray[index];
            if (!item) return undefined;
            return getComparableValue(item);
        };
        var leftIndex = 0;
        var rightIndex = sortedArray.length;
        while (leftIndex <= rightIndex) {
            var middleIndex = Math.floor((leftIndex + rightIndex) / 2);
            var middleValue = getValue(middleIndex);
            if (targetValue < middleValue) {
                rightIndex = middleIndex - 1;
                continue;
            }
            if (targetValue > middleValue) {
                leftIndex = middleIndex + 1;
                continue;
            }
            if (targetValue === middleValue) {
                return {
                    found: true,
                    index: middleIndex
                };
            }
            break;
        }
        return {
            found: false,
            index: leftIndex
        };
    },

    /**
     * Checks if a number is an integer.
     * @param {*} value - The value to check.
     * @returns {boolean} A flag indicating if the value is an integer.
     */
    isInteger: function (value) {
        return value === parseInt(value);
    },

    /**
     * Converts a column number to column name (e.g. 2 -> "B").
     * @param {number} number - The number to convert.
     * @returns {string} The corresponding name.
     */
    columnNumberToName: function (number) {
        if (!this.isInteger(number) || number <= 0) return;

        var dividend = number;
        var name = '';
        var modulo = 0;

        while (dividend > 0) {
            modulo = (dividend - 1) % 26;
            name = String.fromCharCode('A'.charCodeAt(0) + modulo) + name;
            dividend = Math.floor((dividend - modulo) / 26);
        }

        return name;
    },

    /**
     * Converts a column name to column number (e.g. "B" -> 2).
     * @param {string} name - The name to convert.
     * @returns {number} The corresponding number.
     */
    columnNameToNumber: function (name) {
        if (!name || typeof name !== "string") return;

        name = name.toUpperCase();
        var sum = 0;
        for (var i = 0; i < name.length; i++) {
            sum *= 26;
            sum += name[i].charCodeAt(0) - 'A'.charCodeAt(0) + 1;
        }

        return sum;
    },

    /**
     * Converts a row and column (and option sheet) to an Excel address.
     * @param {number} row - The row number.
     * @param {number} column - The column number.
     * @param {string} [sheet] - The sheet name (full a full address).
     * @returns {string} The converted address.
     */
    rowAndColumnToAddress: function (row, column, sheet) {
        if (!this.isInteger(row) || !this.isInteger(column) || row <= 0 || column <= 0) return;
        var address = this.columnNumberToName(column) + row;
        if (sheet) address = this.addressToFullAddress(sheet, address);
        return address;
    },

    /**
     * Converts an address and sheet name to a full address.
     * @param {string} sheet - The sheet name.
     * @param {string} address - The address.
     * @returns {string} The full address.
     */
    addressToFullAddress: function (sheet, address) {
        return "'" + sheet + "'!" + address;
    },

    /**
     * Converts an address to row and column (and sheet if present).
     * @param {string} address - The address to convert.
     * @returns {{}} parsed - The parsed values.
     * @returns {number} parsed.row - The row.
     * @returns {number} parsed.column - The column.
     * @returns {string} [parsed.sheet] - The sheet.
     */
    addressToRowAndColumn: function (address) {
        var match = addressRegex.exec(address);
        if (!match) return;

        var ref = {
            row: parseInt(match[3]),
            column: this.columnNameToNumber(match[2])
        };

        if (match[1]) ref.sheet = match[1];

        return ref;
    },

    /**
     * Converts a date to an Excel number. (The number of days since 1/1/1900, almost.)
     * @param {Date} date - The date to convert
     * @returns {number} - The Excel date number.
     */
    dateToExcelNumber: function (date) {
        var num = (date - dateBase) / millisecondsInDay;

        // "Bug" in Excel that treats 1900 as a leap year.
        if (date > incorrectLeapDate) num += 1;

        return num;
    },

    /**
     * Converts a node to a node string type.
     * @param {xmlDomNode} node - The xpath node.
     * @returns {string} - The xpath node type.
     */
    getNodeType: function (node) {
        return {
            1: 'element',
            2: 'attribute',
            3: 'text',
            8: 'comment',
            9: 'document'
        }[node.nodeType];
    },

    /**
     * Extracts the text from a node.
     * @param {xmlDomNode} node - The xpath node.
     * @returns {string} - The text value of the node.
     **/
    getNodeText: function (node) {
        for (var i = 0; i < node.childNodes.length; i++) {
            if (this.getNodeType(node.childNodes[i]) === 'text') {
                return node.childNodes[i].nodeValue;
            }
        }
        return undefined;
    },

    /**
     * Produce a JSON string describing the node.
     * @param {xmlDomNode} node - The xpath node.
     * @param {object} [info] - Optional dictionary of information to stringify.
     * @returns {string} - The JSON string.
     **/
    getNodeInfo: function (node, info) {
        info = info || {};
        info.nodeName = node.nodeName;
        info.nodeType = this.getNodeType(node);
        info.nodeText = this.getNodeText(node);
        info.nodeValue = node.nodeValue;
        return JSON.stringify(info);
    }
};
