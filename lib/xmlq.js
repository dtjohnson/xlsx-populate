"use strict";

const _ = require("lodash");

module.exports = {
    appendChild(node, child) {
        if (!node.children) node.children = [];
        node.children.push(child);
    },

    setAttributes(node, attributes) {
        _.forOwn(attributes, (value, attribute) => {
            if (_.isNil(value)) {
                if (node.attributes) delete node.attributes[attribute];
            } else {
                if (!node.attributes) node.attributes = {};
                node.attributes[attribute] = value;
            }
        });
    },

    appendChildIfNotFound(node, name) {
        let child = this.findChild(node, name);
        if (!child) {
            child = { name, attributes: {}, children: [] };
            this.appendChild(node, child);
        }

        return child;
    },

    insertBefore(node, child, before) {
        if (!node.children) node.children = [];
        const index = node.children.indexOf(before);
        node.children.splice(index, 0, child);
    },

    isEmpty(node) {
        return _.isEmpty(node.children) && _.isEmpty(node.attributes);
    },

    removeChild(node, child) {
        if (!node.children) return;
        if (typeof child === 'string') {
            _.remove(node.children, { name: child });
        } else {
            const index = node.children.indexOf(child);
            node.children.splice(index, 1);
        }
    },

    hasChild(node, name) {
        return _.some(node.children, { name });
    },

    findChild(node, name) {
        return _.find(node.children, { name });
    },

    getChildAttribute(node, name, attribute) {
        const child = this.findChild(node, name);
        if (child) return child.attributes && child.attributes[attribute];
    },

    setChildAttributes(node, name, attributes) {
        let child = this.findChild(node, name);
        _.forOwn(attributes, (value, attribute) => {
            if (_.isNil(value)) {
                if (child && child.attributes) delete child.attributes[attribute];
            } else {
                if (!child) {
                    child = { name, attributes: {}, children: [] };
                    this.appendChild(node, child);
                }

                if (!child.attributes) child.attributes = {};
                child.attributes[attribute] = value;
            }
        });

        return child;
    },

    removeChildIfEmpty(node, child) {
        if (typeof child === 'string') child = this.findChild(node, child);
        if (child && this.isEmpty(child)) this.removeChild(node, child);
    }
};
