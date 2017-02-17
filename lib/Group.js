"use strict";

// TODO: tests
// TODO: JSDoc

const ArgHandler = require("./ArgHandler");
const debug = require("./debug")('Group');

class Group {
    constructor(items) {
        this._items = items;
    }

    forEach(handler) {
        this._items.forEach(handler);
        return this;
    }

    map(handler) {
        return this._items.map(handler);
    }

    // TODO: Is this necessary?
    forEachValue(values, handler) {
        return this.forEach((item, i) => {
            if (values[i] !== undefined) handler(values[i], item, i);
        });
    }

    reduce() {
        // TODO
    }

    item(i) {
        return this._items[i];
    }

    formula() {
        // TODO: What to do here?
    }

    value() {
        debug("value(%o)", arguments);
        return new ArgHandler("Group.value")
            .case(() => {
                // Get values
                return this.map(cell => cell.value());
            })
            .case('function', valueFunc => {
                // Set a value for the items to the result of a function
                return this.forEach((item, i) => {
                    item.value(valueFunc(item, i));
                });
            })
            .case('array', values => {
                // Set value for the items using an array of matching dimension
                return this.forEachValue(values, (value, item) => item.value(value));
            })
            .case('*', value => {
                // Set the value for all cells to a single value
                return this.forEach(item => item.value(value));
            })
            .handle(arguments);
    }

    style() {
        debug("style(%o)", arguments);
        return new ArgHandler("Cell.style")
            .case('string', name => {
                // Get single value
                return this.map(item => item.style(name));
            })
            .case('array', names => {
                // Get list of values
                const values = {};
                names.forEach(name => {
                    values[name] = this.style(name);
                });

                return values;
            })
            .case(['string', 'function'], (name, valueFunc) => {
                // Set a single value for the items to the result of a function
                return this.forEach((item, i) => {
                    item.style(name, valueFunc(item, i));
                });
            })
            .case(['string', 'array'], (name, values) => {
                // Set a single value for the items using an array of matching dimension
                return this.forEachValue(values, (value, item) => item.style(name, value));
            })
            .case(['string', '*'], (name, value) => {
                // Set a single value for all items to a single value
                return this.forEach(item => item.style(name, value));
            })
            .case('object', nameValues => {
                // Object of key value pairs to set
                for (const name in nameValues) {
                    if (!nameValues.hasOwnProperty(name)) continue;
                    const value = nameValues[name];
                    this.style(name, value);
                }

                return this;
            })
            .handle(arguments);
    }
}

module.exports = Group;
