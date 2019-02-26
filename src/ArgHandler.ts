/**
 * TODO: Deprecated
 * @module xlsx-populate
 */

import * as _ from 'lodash';

export type Handler = (...args: any[]) => any;
export interface ICase {
    types: string[];
    handler: Handler;
}

/**
 * Method argument handler. Used for overloading methods.
 * @hidden
 */
export class ArgHandler {
    private _name: string;
    private _cases: ICase[] = [];

    /**
     * Creates a new instance of ArgHandler.
     * @param name - The method name to use in error messages.
     */
    public constructor(name: string) {
        this._name = name;
        this._cases = [];
    }

    /**
     * Add a case.
     * @param types - The type or types of arguments to match this case.
     * @param handler - The function to call when this case is matched.
     * @returns The handler for chaining.
     * TODO: Proper TSDoc for overloads
     */
    public case(types: string|string[], handler: Handler): this;
    public case(handler: Handler): this;
    public case(typesOrHandler: string|string[]|Handler, optHandler?: Handler): this {
        let types: string[];
        let handler: Handler;

        if (typeof typesOrHandler === 'function') {
            types = [];
            handler = typesOrHandler;
        } else if (optHandler) {
            types = Array.isArray(typesOrHandler) ? typesOrHandler : [ typesOrHandler ];
            handler = optHandler;
        } else {
            throw new Error('Invalid case.');
        }

        this._cases.push({ types, handler });
        return this;
    }

    /**
     * Handle the method arguments by checking each case in order until one matches and then call its handler.
     * @param args - The method arguments.
     * @returns The result of the handler.
     * @throws {Error} Throws if no case matches.
     */
    public handle(args: any[]): any {
        for (let i = 0; i < this._cases.length; i++) {
            const c = this._cases[i];
            if (this._argsMatchTypes(args, c.types)) {
                return c.handler.bind(undefined)(...args);
            }
        }

        throw new Error(`${this._name}: Invalid arguments.`);
    }

    /**
     * Check if the arguments match the given types.
     * @param args - The arguments.
     * @param types - The types.
     * @returns True if matches, false otherwise.
     * @throws if unknown type.
     */
    private _argsMatchTypes(args: any[], types: string[]): boolean {
        if (args.length !== types.length) return false;

        return _.every(args, (arg, i) => {
            const type = types[i];

            if (type === '*') return true;
            if (type === 'nil') return _.isNil(arg);
            if (type === 'string') return typeof arg === 'string';
            if (type === 'boolean') return typeof arg === 'boolean';
            if (type === 'number') return typeof arg === 'number';
            if (type === 'integer') return typeof arg === 'number' && _.isInteger(arg);
            if (type === 'function') return typeof arg === 'function';
            if (type === 'array') return Array.isArray(arg);
            if (type === 'date') return arg && arg.constructor === Date;
            if (type === 'object') return arg && arg.constructor === Object;
            if (arg && arg.constructor && arg.constructor.name === type) return true;

            throw new Error(`Unknown type: ${type}`);
        });
    }
}
