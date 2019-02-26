"use strict";
/**
 * @module xlsx-populate
 */
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (Object.hasOwnProperty.call(mod, k)) result[k] = mod[k];
    result["default"] = mod;
    return result;
};
Object.defineProperty(exports, "__esModule", { value: true });
const _ = __importStar(require("lodash"));
/**
 * Method argument handler. Used for overloading methods.
 * @hidden
 */
class ArgHandler {
    /**
     * Creates a new instance of ArgHandler.
     * @param name - The method name to use in error messages.
     */
    constructor(name) {
        this._cases = [];
        this._name = name;
        this._cases = [];
    }
    case(typesOrHandler, optHandler) {
        let types;
        let handler;
        if (typeof typesOrHandler === 'function') {
            types = [];
            handler = typesOrHandler;
        }
        else if (optHandler) {
            types = Array.isArray(typesOrHandler) ? typesOrHandler : [typesOrHandler];
            handler = optHandler;
        }
        else {
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
    handle(args) {
        for (let i = 0; i < this._cases.length; i++) {
            const c = this._cases[i];
            if (this._argsMatchTypes(args, c.types)) {
                return c.handler.apply(undefined, args);
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
    _argsMatchTypes(args, types) {
        if (args.length !== types.length)
            return false;
        return _.every(args, (arg, i) => {
            const type = types[i];
            if (type === '*')
                return true;
            if (type === 'nil')
                return _.isNil(arg);
            if (type === 'string')
                return typeof arg === 'string';
            if (type === 'boolean')
                return typeof arg === 'boolean';
            if (type === 'number')
                return typeof arg === 'number';
            if (type === 'integer')
                return typeof arg === 'number' && _.isInteger(arg);
            if (type === 'function')
                return typeof arg === 'function';
            if (type === 'array')
                return Array.isArray(arg);
            if (type === 'date')
                return arg && arg.constructor === Date;
            if (type === 'object')
                return arg && arg.constructor === Object;
            if (arg && arg.constructor && arg.constructor.name === type)
                return true;
            throw new Error(`Unknown type: ${type}`);
        });
    }
}
exports.ArgHandler = ArgHandler;
//# sourceMappingURL=ArgHandler.js.map