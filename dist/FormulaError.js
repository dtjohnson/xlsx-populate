"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
/**
 * A formula error (e.g. #DIV/0!).
 */
class FormulaError {
    // /**
    //  * Creates a new instance of Formula Error.
    //  * @param {string} error - The error code.
    //  */
    constructor(_error) {
        this._error = _error;
    }
    /**
     * Get the matching FormulaError object.
     * @param {string} error - The error code.
     * @returns {FormulaError} The matching FormulaError or a new object if no match.
     * @ignore
     */
    static getError(error) {
        return FormulaError.ERRORS.find(value => {
            return value instanceof FormulaError && value.error() === error;
        }) || new FormulaError(error);
    }
    ;
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
FormulaError.ERRORS = [
    FormulaError.DIV0,
    FormulaError.NA,
    FormulaError.NAME,
    FormulaError.NULL,
    FormulaError.NUM,
    FormulaError.REF,
    FormulaError.VALUE
];
exports.FormulaError = FormulaError;
//# sourceMappingURL=FormulaError.js.map