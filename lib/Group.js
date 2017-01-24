"use strict";

// TODO: tests
// TODO: JSDoc

const _ArgParser = require("./_ArgParser");
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

    fullAddress() {
        // TODO
    }

    formula() {
        // TODO: What to do here?
    }

    value() {
        debug("value(%o)", arguments);
        return new _ArgParser("Group.value")
            .case(() => {
                // Get values
                return this.map(cell => cell.value());
            })
            .case(Function, valueFunc => {
                // Set a value for the items to the result of a function
                return this.forEach((item, i) => {
                    item.value(valueFunc(item, i));
                });
            })
            .case(Array, values => {
                // Set value for the items using an array of matching dimension
                return this.forEachValue(values, (value, item) => item.value(value));
            })
            .case(undefined, value => {
                // Set the value for all cells to a single value
                return this.forEach(item => item.value(value));
            })
            .parse(arguments);
    }

    style() {
        debug("style(%o)", arguments);
        return new _ArgParser("Cell.style")
            .case(String, name => {
                // Get single value
                return this.map(item => item.style(name));
            })
            .case(Array, names => {
                // Get list of values
                const values = {};
                names.forEach(name => {
                    values[name] = this.style(name);
                });

                return values;
            })
            .case([String, Function], (name, valueFunc) => {
                // Set a single value for the items to the result of a function
                return this.forEach((item, i) => {
                    item.style(name, valueFunc(item, i));
                });
            })
            .case([String, Array], (name, values) => {
                // Set a single value for the items using an array of matching dimension
                return this.forEachValue(values, (value, item) => item.style(name, value));
            })
            .case([String, undefined], (name, value) => {
                // Set a single value for all items to a single value
                return this.forEach(item => item.style(name, value));
            })
            .case(Object, nameValues => {
                // Object of key value pairs to set
                for (const name in nameValues) {
                    if (!nameValues.hasOwnProperty(name)) continue;
                    const value = nameValues[name];
                    this.style(name, value);
                }

                return this;
            })
            .parse(arguments);
    }
}

module.exports = Group;
