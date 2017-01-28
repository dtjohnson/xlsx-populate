"use strict";

// TODO: Tests
// TODO: JSDoc

const debug = require("./debug")('xq');
const utils = require("./utils");

module.exports = {
    query(node, query, projection) {
        debug("query(?, %o)", query);
        let result = this._query(node, query, {});
        if (!result || !projection) return result;

        if (projection === true) projection = "$last";
        const parts = projection.split(".");
        for (let i = 0; i < parts.length; i++) {
            const part = parts[i];
            result = result[part];
            if (result === undefined) return null;
        }

        return result;
    },

    _query(node, query, result, fullResult) {
        debug("_query(?, %o, %o)", query, result);
        if (!fullResult) fullResult = result;
        for (const key in query) {
            if (!query.hasOwnProperty(key)) return;
            const queryValue = query[key];

            if (key === "#text") {
                const textValue = node.textContent;
                const value = this.convertFromString(textValue, queryValue);
                if (value === null) return null;
                result[key] = fullResult.$last = value;
            } else if (key[0] === "@") {
                const name = key.substr(1);
                if (node.hasAttribute(name)) {
                    if (queryValue === null) return null;
                    const textValue = node.getAttribute(name);
                    const value = this.convertFromString(textValue, queryValue.$type || queryValue);
                    if (value === null) return null;
                    result[key] = fullResult.$last = value;
                } else if (queryValue && !queryValue.$optional) {
                    return null;
                }
            } else if (key[0] !== "$") {
                const childNodes = node.getElementsByTagName(key);
                result[key] = [];
                for (let i = 0; i < childNodes.length; i++) {
                    const childNode = childNodes[i];
                    if (queryValue === null) return null;
                    const childResult = this._query(childNode, queryValue, {}, fullResult);
                    if (childResult === null) return null;
                    result[key].push(childResult);
                }

                if (!result[key].length && queryValue && !queryValue.$optional) return null;
                if (!queryValue.$multi) result[key] = result[key][0];
            }
        }

        return result;
    },

    convertToString(value) {
        debug("convertToString(%o)", arguments);
        if (value === null || value === undefined) return "";
        if (typeof value === "boolean") return value ? "1" : "0";
        if (typeof value === "number") return value.toString();
        if (value instanceof Date) return this.convertFromString(utils.dateToExcelNumber(value));
        return value;
    },

    convertFromString(textValue, queryValue) {
        debug("convertFromString(%o)", arguments);
        if (queryValue === Boolean) return textValue === "1";
        if (queryValue === Number) return parseFloat(textValue);
        if (queryValue === String) return textValue;
        if (queryValue === Date) return utils.dateToExcelNumber(this.convertFromString(textValue, Number));
        if (typeof queryValue === "number") {
            return queryValue === this.convertFromString(textValue, Number) ? queryValue : null;
        }
        if (typeof queryValue === "boolean") {
            return queryValue === this.convertFromString(textValue, Boolean) ? queryValue : null;
        }
        if (typeof queryValue === "string") {
            return queryValue === textValue ? queryValue : null;
        }

        return null;
    },

    update(node, update) {
        debug("update(?, %o)", update);
        for (const key in update) {
            if (!update.hasOwnProperty(key)) return;

            const value = update[key];
            if (key === "#text") {
                node.textContent = this.convertToString(value);
            } else if (key[0] === "@") {
                // Attribute
                const attributeName = key.substr(1);
                if (value === null || value === undefined) {
                    node.removeAttribute(attributeName);
                } else {
                    const textValue = this.convertToString(value);
                    node.setAttribute(attributeName, textValue);
                }
            } else if (key[0] !== "$") {
                // Child node
                let childNode = node.getElementsByTagName(key)[0];
                if (value === null || value === undefined) {
                    if (childNode) node.removeChild(childNode);
                } else {
                    if (!childNode || value.$append) {
                        childNode = node.ownerDocument.createElement(key);
                        node.appendChild(childNode);
                    }

                    this.update(childNode, value);
                }
            }
        }

        if (update.$removeIfEmpty) {
            if (!node.hasAttributes() && node.childNodes.length === 0) {
                node.parentNode.removeChild(node);
            }
        }
    }
};
