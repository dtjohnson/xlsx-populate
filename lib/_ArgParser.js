"use strict";

// TODO: Tests
// TODO: JSDoc
// TODO: Debugs

class _ArgParser {
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
        for (let i = 0; i < types.length; i++) {
            const type = types[i];
            const arg = args[i];
            if (type === String && typeof arg === "string") continue;
            if (type === Boolean && typeof arg === "boolean") continue;
            if (type === Number && typeof arg === "number") continue;
            if (type === Function && typeof arg === "function") continue;
            if (type === Array && Array.isArray(arg)) continue;
            if (type === Date && arg && arg.constructor === Date) continue;
            if (type === Object && arg && arg.constructor === Object) continue;
            if (type !== undefined && type === arg) continue; // TODO: Do we have any actually need for specific values?
            if (type === undefined) continue;
            return false;
        }

        return true;
    }

    parse(args) {
        for (let i = 0; i < this._cases.length; i++) {
            const c = this._cases[i];
            if (this._argsMatchTypes(args, c.types)) {
                return c.handler.apply(null, args);
            }
        }

        throw new Error(`${this._name}: Invalid arguments.`);
    }
}

module.exports = _ArgParser;
