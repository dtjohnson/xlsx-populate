/**
 * A formula error (e.g. #DIV/0!).
 */
export declare class FormulaError {
    private _error;
    /**
     * \#DIV/0! error.
     * @type {FormulaError}
     */
    static DIV0: FormulaError;
    /**
     * \#N/A error.
     * @type {FormulaError}
     */
    static NA: FormulaError;
    /**
     * \#NAME? error.
     * @type {FormulaError}
     */
    static NAME: FormulaError;
    /**
     * \#NULL! error.
     * @type {FormulaError}
     */
    static NULL: FormulaError;
    /**
     * \#NUM! error.
     * @type {FormulaError}
     */
    static NUM: FormulaError;
    /**
     * \#REF! error.
     * @type {FormulaError}
     */
    static REF: FormulaError;
    /**
     * \#VALUE! error.
     * @type {FormulaError}
     */
    static VALUE: FormulaError;
    static ERRORS: FormulaError[];
    /**
     * Get the matching FormulaError object.
     * @param {string} error - The error code.
     * @returns {FormulaError} The matching FormulaError or a new object if no match.
     * @ignore
     */
    static getError(error: string): FormulaError;
    constructor(_error: string);
    /**
     * Get the error code.
     * @returns {string} The error code.
     */
    error(): string;
}
//# sourceMappingURL=FormulaError.d.ts.map