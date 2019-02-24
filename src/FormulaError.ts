/**
 * @module xlsx-populate
 */

/**
 * A formula error (e.g. #DIV/0!).
 */
export class FormulaError {
    /**
     * \#DIV/0! error.
     */
    public static readonly DIV0 = new FormulaError('#DIV/0!');

    /**
     * \#N/A error.
     */
    public static readonly NA = new FormulaError('#N/A');

    /**
     * \#NAME? error.
     */
    public static readonly NAME = new FormulaError('#NAME?');

    /**
     * \#NULL! error.
     */
    public static readonly NULL = new FormulaError('#NULL!');

    /**
     * \#NUM! error.
     */
    public static readonly NUM = new FormulaError('#NUM!');

    /**
     * \#REF! error.
     */
    public static readonly REF = new FormulaError('#REF!');

    /**
     * \#VALUE! error.
     */
    public static readonly VALUE = new FormulaError('#VALUE!');

    /**
     * Array of standard errors.
     */
    public static readonly ERRORS: ReadonlyArray<FormulaError> = [
        FormulaError.DIV0,
        FormulaError.NA,
        FormulaError.NAME,
        FormulaError.NULL,
        FormulaError.NUM,
        FormulaError.REF,
        FormulaError.VALUE,
    ];

    /**
     * Get the matching FormulaError object.
     * @param error - The error code.
     * @returns The matching FormulaError or a new object if no match.
     * @ignore
     */
    public static getError(error: string): FormulaError {
        return FormulaError.ERRORS.find(value => {
            return value instanceof FormulaError && value.error() === error;
        }) || new FormulaError(error);
    }

    /**
     * Creates a new instance of Formula Error.
     * @param error - The error code.
     */
    private constructor(private readonly _error: string) {
    }

    /**
     * Get the error code.
     * @returns The error code.
     */
    public error(): string {
        return this._error;
    }
}
