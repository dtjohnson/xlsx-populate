"use strict";

const _ = require("lodash");

module.exports = {
    has(object, path) {
        return _.has(object, path);
    },

    apply(object, path, func) {
        return func(_.get(object, path));
    },

    get(object, path) {
        return _.get(object, path);
    },

    set() {
        if (arguments.length === 2) {
            const object = arguments[0], sets = arguments[1];
            _.forOwn(sets, (value, path) => this.set(object, path, value));
            return object;
        } else {
            const object = arguments[0], path = arguments[1], value = arguments[2];
            if (value === null || value === undefined) {
                _.unset(object, path);
            } else {
                _.set(object, path, value);
            }

            return object;
        }
    },

    setIfNeeded(object, path, value) {
        if (!this.has(object, path)) this.set(object, path, value);
    },

    isEmpty(object, path) {
        return _.isEmpty(_.get(object, path));
    }
};
