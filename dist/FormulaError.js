"use strict";
/**
 * @module xlsx-populate
 */
Object.defineProperty(exports, "__esModule", { value: true });
/**
 * A formula error (e.g. #DIV/0!).
 */
class FormulaError {
    /**
     * Creates a new instance of Formula Error.
     * @param error - The error code.
     */
    constructor(_error) {
        this._error = _error;
    }
    /**
     * Get the matching FormulaError object.
     * @param error - The error code.
     * @returns The matching FormulaError or a new object if no match.
     * @ignore
     */
    static getError(error) {
        return FormulaError.ERRORS.find(value => {
            return value instanceof FormulaError && value.error() === error;
        }) || new FormulaError(error);
    }
    /**
     * Get the error code.
     * @returns The error code.
     */
    error() {
        return this._error;
    }
}
/**
 * \#DIV/0! error.
 */
FormulaError.DIV0 = new FormulaError('#DIV/0!');
/**
 * \#N/A error.
 */
FormulaError.NA = new FormulaError('#N/A');
/**
 * \#NAME? error.
 */
FormulaError.NAME = new FormulaError('#NAME?');
/**
 * \#NULL! error.
 */
FormulaError.NULL = new FormulaError('#NULL!');
/**
 * \#NUM! error.
 */
FormulaError.NUM = new FormulaError('#NUM!');
/**
 * \#REF! error.
 */
FormulaError.REF = new FormulaError('#REF!');
/**
 * \#VALUE! error.
 */
FormulaError.VALUE = new FormulaError('#VALUE!');
/**
 * Array of standard errors.
 */
FormulaError.ERRORS = [
    FormulaError.DIV0,
    FormulaError.NA,
    FormulaError.NAME,
    FormulaError.NULL,
    FormulaError.NUM,
    FormulaError.REF,
    FormulaError.VALUE,
];
exports.FormulaError = FormulaError;
//# sourceMappingURL=FormulaError.js.map