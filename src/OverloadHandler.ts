/**
 * @ignore
 */


/**
 * Type selector
 */
type TypeSelector = undefined|string|(new (...args: any[]) => any);

/**
 * Interface for a overload handling case.
 */
interface ICase {
    typeSelectors: TypeSelector[];
    handler(...args: any[]): any;
}

/**
 * Handler used to simplify function overloading.
 * @hidden
 */
export class OverloadHandler {
    /**
     * Array of cases.
     */
    private cases: ICase[] = [];

    /**
     * Creates a new instance of OverloadHandler
     * @param name - Name to use in error messages.
     */
    public constructor(private name: string) {}

    /**
     * Add a case with no arguments and no return value.
     * @param handler - Function to handle the case.
     */
    public case(
        handler: () => void,
    ): this;
    /**
     * Add a case with no arguments and a return value.
     * @param handler - Function to handle the case.
     */
    public case<TResult>(
        handler: () => TResult,
    ): this;
    /**
     * Add a case with 1 argument and a return value.
     * @param typeSelector - Selector for the argument.
     * @param handler - Function to handle the case.
     */
    public case<TArg, TResult>(
        typeSelector: TypeSelector,
        handler: (arg: TArg) => TResult,
    ): this;
    /**
     * Add a case with 2 arguments and a return value.
     * @param typeSelector1 - Selector for the first argument.
     * @param typeSelector2 - Selector for the second argument.
     * @param handler - Function to handle the case.
     */
    public case<TArg1, TArg2, TResult>(
        typeSelector1: TypeSelector,
        typeSelector2: TypeSelector,
        handler: (arg1: TArg1, arg2: TArg2) => TResult,
    ): this;
    /**
     * Add a case with 3 arguments and a return value.
     * @param typeSelector1 - Selector for the first argument.
     * @param typeSelector2 - Selector for the second argument.
     * @param typeSelector3 - Selector for the third argument.
     * @param handler - Function to handle the case.
     */
    public case<TArg1, TArg2, TArg3, TResult>(
        typeSelector1: TypeSelector,
        typeSelector2: TypeSelector,
        typeSelector3: TypeSelector,
        handler: (arg1: TArg1, arg2: TArg2, arg3: TArg3) => TResult,
    ): this;
    /**
     * Add a case with 4 arguments and a return value.
     * @param typeSelector1 - Selector for the first argument.
     * @param typeSelector2 - Selector for the second argument.
     * @param typeSelector3 - Selector for the third argument.
     * @param typeSelector4 - Selector for the fourth argument.
     * @param handler - Function to handle the case.
     */
    public case<TArg1, TArg2, TArg3, TArg4, TResult>(
        typeSelector1: TypeSelector,
        typeSelector2: TypeSelector,
        typeSelector3: TypeSelector,
        typeSelector4: TypeSelector,
        handler: (arg1: TArg1, arg2: TArg2, arg3: TArg3, arg4: TArg4) => TResult,
    ): this;
    public case(...args: any[]): this {
        const larg = args.length - 1;
        this.cases.push({
            typeSelectors: args.slice(0, larg),
            handler: args[larg],
        });

        return this;
    }

    /**
     * Handle arguments from a function call.
     * @param args - The arguments.
     */
    public handle(args: any[]): any {
        for (let i = 0; i < this.cases.length; i++) {
            const c = this.cases[i];
            if (this.argsMatchTypeSelectors(args, c.typeSelectors)) {
                return c.handler.bind(undefined)(...args);
            }
        }

        throw new Error(`${this.name}: Invalid arguments.`);
    }

    /**
     * Check if the arguments match the type selectors.
     * @param args - The arguments.
     * @param typeSelectors - The type selectors.
     */
    private argsMatchTypeSelectors(args: any[], typeSelectors: TypeSelector[]): boolean {
        if (args.length !== typeSelectors.length) return false;

        return args.every((arg, i) => {
            const typeSelector = typeSelectors[i];

            if (typeSelector === 'any') return true;
            if (typeSelector === undefined) return arg === null || arg === undefined;
            if (typeof typeSelector === 'string') return typeof arg === typeSelector;
            return arg instanceof typeSelector;
        });
    }
}
