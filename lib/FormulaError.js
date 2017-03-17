"use strict";

const _ = require("lodash");

/**
 * A formula error (e.g. #DIV/0!).
 */
class FormulaError {
    // /**
    //  * Creates a new instance of Formula Error.
    //  * @param {string} error - The error code.
    //  */
    constructor(error) {
        this._error = error;
    }

    /**
     * Get the error code.
     * @returns {string} The error code.
     */
    error() {
        return this._error;
    }
}

/**
 * \#DIV/0! error.
 * @type {FormulaError}
 */
FormulaError.DIV0 = new FormulaError("#DIV/0!");

/**
 * \#N/A error.
 * @type {FormulaError}
 */
FormulaError.NA = new FormulaError("#N/A");

/**
 * \#NAME? error.
 * @type {FormulaError}
 */
FormulaError.NAME = new FormulaError("#NAME?");

/**
 * \#NULL! error.
 * @type {FormulaError}
 */
FormulaError.NULL = new FormulaError("#NULL!");

/**
 * \#NUM! error.
 * @type {FormulaError}
 */
FormulaError.NUM = new FormulaError("#NUM!");

/**
 * \#REF! error.
 * @type {FormulaError}
 */
FormulaError.REF = new FormulaError("#REF!");

/**
 * \#VALUE! error.
 * @type {FormulaError}
 */
FormulaError.VALUE = new FormulaError("#VALUE!");

/**
 * Get the matching FormulaError object.
 * @param {string} error - The error code.
 * @returns {FormulaError} The matching FormulaError or a new object if no match.
 * @ignore
 */
FormulaError.getError = error => {
    return _.find(FormulaError, value => {
        return value instanceof FormulaError && value.error() === error;
    }) || new FormulaError(error);
};

module.exports = FormulaError;
