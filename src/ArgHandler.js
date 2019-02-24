"use strict";

const _ = require("lodash");

/**
 * Method argument handler. Used for overloading methods.
 * @private
 */
class ArgHandler {
    /**
     * Creates a new instance of ArgHandler.
     * @param {string} name - The method name to use in error messages.
     */
    constructor(name) {
        this._name = name;
        this._cases = [];
    }

    /**
     * Add a case.
     * @param {string|Array.<string>} [types] - The type or types of arguments to match this case.
     * @param {Function} handler - The function to call when this case is matched.
     * @returns {ArgHandler} The handler for chaining.
     */
    case(types, handler) {
        if (arguments.length === 1) {
            handler = types;
            types = [];
        }

        if (!Array.isArray(types)) types = [types];
        this._cases.push({ types, handler });
        return this;
    }

    /**
     * Handle the method arguments by checking each case in order until one matches and then call its handler.
     * @param {Arguments|Array.<*>} args - The method arguments.
     * @returns {*} The result of the handler.
     * @throws {Error} Throws if no case matches.
     */
    handle(args) {
        for (let i = 0; i < this._cases.length; i++) {
            const c = this._cases[i];
            if (this._argsMatchTypes(args, c.types)) {
                return c.handler.apply(null, args);
            }
        }

        throw new Error(`${this._name}: Invalid arguments.`);
    }

    /**
     * Check if the arguments match the given types.
     * @param {Arguments} args - The arguments.
     * @param {Array.<string>} types - The types.
     * @returns {boolean} True if matches, false otherwise.
     * @throws {Error} Throws if unknown type.
     * @private
     */
    _argsMatchTypes(args, types) {
        if (args.length !== types.length) return false;

        return _.every(args, (arg, i) => {
            const type = types[i];

            if (type === '*') return true;
            if (type === 'nil') return _.isNil(arg);
            if (type === 'string') return typeof arg === "string";
            if (type === 'boolean') return typeof arg === "boolean";
            if (type === 'number') return typeof arg === "number";
            if (type === 'integer') return typeof arg === "number" && _.isInteger(arg);
            if (type === 'function') return typeof arg === "function";
            if (type === 'array') return Array.isArray(arg);
            if (type === 'date') return arg && arg.constructor === Date;
            if (type === 'object') return arg && arg.constructor === Object;
            if (arg && arg.constructor && arg.constructor.name === type) return true;

            throw new Error(`Unknown type: ${type}`);
        });
    }
}

module.exports = ArgHandler;
