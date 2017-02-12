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
        if (!this.has(object, path)) return this.set(object, path, value);
    },

    findAndSet(array, predicate, additionalValues) {
        if (!_.some(array, predicate)) array.push(predicate);
        const item = _.find(array, predicate);
        if (additionalValues) _.forOwn(additionalValues, (value, path) => _.set(item, path, value));
        return item;
    },

    isEmpty(object, path) {
        return _.isEmpty(_.get(object, path));
    },

    insertBefore(array, item, before) {
        const index = array.indexOf(before);
        array.splice(index, 0, item);
        return array;
    },

    findAndGet(array, predicate, getPath) {
        const item = _.find(array, predicate);
        return _.get(item, getPath);
    }
};
