/**
 * A formula error (e.g. #DIV/0!).
 */
export class FormulaError {
    /**
     * \#DIV/0! error.
     * @type {FormulaError}
     */
    static DIV0 = new FormulaError("#DIV/0!");

    /**
     * \#N/A error.
     * @type {FormulaError}
     */
    static NA = new FormulaError("#N/A");

    /**
     * \#NAME? error.
     * @type {FormulaError}
     */
    static NAME = new FormulaError("#NAME?");

    /**
     * \#NULL! error.
     * @type {FormulaError}
     */
    static NULL = new FormulaError("#NULL!");

    /**
     * \#NUM! error.
     * @type {FormulaError}
     */
    static NUM = new FormulaError("#NUM!");

    /**
     * \#REF! error.
     * @type {FormulaError}
     */
    static REF = new FormulaError("#REF!");

    /**
     * \#VALUE! error.
     * @type {FormulaError}
     */
    static VALUE = new FormulaError("#VALUE!");

    static ERRORS: FormulaError[] = [
        FormulaError.DIV0,
        FormulaError.NA,
        FormulaError.NAME,
        FormulaError.NULL,
        FormulaError.NUM,
        FormulaError.REF,
        FormulaError.VALUE
    ];

    /**
     * Get the matching FormulaError object.
     * @param {string} error - The error code.
     * @returns {FormulaError} The matching FormulaError or a new object if no match.
     * @ignore
     */
    static getError(error: string): FormulaError {
        return FormulaError.ERRORS.find(value => {
            return value instanceof FormulaError && value.error() === error;
        }) || new FormulaError(error);
    };

    // /**
    //  * Creates a new instance of Formula Error.
    //  * @param {string} error - The error code.
    //  */
    constructor(private _error: string) {
    }

    /**
     * Get the error code.
     * @returns {string} The error code.
     */
    error() {
        return this._error;
    }
}
