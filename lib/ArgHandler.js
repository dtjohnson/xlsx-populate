"use strict";

// TODO: Tests
// TODO: JSDoc
// TODO: Debugs

const _ = require("lodash");

class ArgHandler {
    constructor(name) {
        this._name = name;
        this._cases = [];
    }

    case(types, handler) {
        if (arguments.length === 1) {
            handler = types;
            types = [];
        }

        if (!Array.isArray(types)) types = [types];
        this._cases.push({ types, handler });
        return this;
    }

    // TODO: Do we need to support allowed values (e.g. alignment can only be 'top', 'center', and 'bottom')?
    // TODO: Do we need a function callback support?
    _argsMatchTypes(args, types) {
        if (args.length !== types.length) return false;

        return _.every(args, (arg, i) => {
            const type = types[i];

            if (type === '*') return true;
            if (type === 'nil') return _.isNil(arg);
            if (type === 'string') return typeof arg === "string";
            if (type === 'boolean') return typeof arg === "boolean";
            if (type === 'number') return typeof arg === "number";
            if (type === 'integer') return typeof arg === "number" && Number.isInteger(arg);
            if (type === 'function') return typeof arg === "function";
            if (type === 'array') return Array.isArray(arg);
            if (type === 'date') return arg && arg.constructor === Date;
            if (type === 'object') return arg && arg.constructor === Object;

            throw new Error(`Unknown type: ${type}`);
        });
    }

    handle(args) {
        for (let i = 0; i < this._cases.length; i++) {
            const c = this._cases[i];
            if (this._argsMatchTypes(args, c.types)) {
                return c.handler.apply(null, args);
            }
        }

        throw new Error(`${this._name}: Invalid arguments.`);
    }
}

module.exports = ArgHandler;
