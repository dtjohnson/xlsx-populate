"use strict";
/**
 * @module xlsx-populate
 */
Object.defineProperty(exports, "__esModule", { value: true });
/**
 * Handler used to simplify function overloading.
 * @hidden
 */
class OverloadHandler {
    /**
     * Creates a new instance of OverloadHandler
     * @param name - Name to use in error messages.
     */
    constructor(name) {
        this.name = name;
        /**
         * Array of cases.
         */
        this.cases = [];
    }
    case(...args) {
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
    handle(args) {
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
    argsMatchTypeSelectors(args, typeSelectors) {
        if (args.length !== typeSelectors.length)
            return false;
        return args.every((arg, i) => {
            const typeSelector = typeSelectors[i];
            if (typeSelector === 'any')
                return true;
            if (typeSelector === undefined)
                return arg === null || arg === undefined;
            if (typeof typeSelector === 'string')
                return typeof arg === typeSelector;
            return arg instanceof typeSelector;
        });
    }
}
exports.OverloadHandler = OverloadHandler;
//# sourceMappingURL=OverloadHandler.js.map