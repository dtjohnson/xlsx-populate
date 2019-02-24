/**
 * @module xlsx-populate
 */
/**
 * A formula error (e.g. #DIV/0!).
 */
export declare class FormulaError {
    private readonly _error;
    /**
     * \#DIV/0! error.
     */
    static readonly DIV0: FormulaError;
    /**
     * \#N/A error.
     */
    static readonly NA: FormulaError;
    /**
     * \#NAME? error.
     */
    static readonly NAME: FormulaError;
    /**
     * \#NULL! error.
     */
    static readonly NULL: FormulaError;
    /**
     * \#NUM! error.
     */
    static readonly NUM: FormulaError;
    /**
     * \#REF! error.
     */
    static readonly REF: FormulaError;
    /**
     * \#VALUE! error.
     */
    static readonly VALUE: FormulaError;
    /**
     * Array of standard errors.
     */
    static readonly ERRORS: ReadonlyArray<FormulaError>;
    /**
     * Get the matching FormulaError object.
     * @param error - The error code.
     * @returns The matching FormulaError or a new object if no match.
     * @ignore
     */
    static getError(error: string): FormulaError;
    /**
     * Creates a new instance of Formula Error.
     * @param error - The error code.
     */
    private constructor();
    /**
     * Get the error code.
     * @returns The error code.
     */
    error(): string;
}
//# sourceMappingURL=FormulaError.d.ts.map