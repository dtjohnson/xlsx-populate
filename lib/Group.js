"use strict";

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

    forEachValue(values, handler) {
        return this.forEach((item, i) => {
            if (values[i] !== undefined) handler(values[i], item, i);
        });
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
                this.forEach((item, i) => {
                    item.style(name, valueFunc(item, i));
                });

                return this;
            })
            .case([String, Array], (name, values) => {
                // Set a single value for the items using an array of matching dimension
                this.forEachValue(values, (value, item) => item.style(name, value));
                return this;
            })
            .case([String, undefined], (name, value) => {
                // Set a single value for all items to a single value
                this.forEach(item => item.style(name, value));
                return this;
            })
            .case(Object, nameValues => {
                // Object of key value pairs to set
                for (const name in nameValues) {
                    if (!nameValues.hasOwnProperty(name)) continue;
                    const value = nameValues[name];
                    this.style(name, value);
                }
            })
            .parse(arguments);
    }
}

module.exports = Group;
