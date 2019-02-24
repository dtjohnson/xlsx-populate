"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const _ = require("lodash");
/**
 * Append a child to the node.
 * @param {{}} node - The parent node.
 * @param {{}} child - The child node.
 * @returns {undefined}
 */
function appendChild(node, child) {
    if (!node.children)
        node.children = [];
    node.children.push(child);
}
exports.appendChild = appendChild;
/**
 * Append a child if one with the given name is not found.
 * @param {{}} node - The parent node.
 * @param {string} name - The child node name.
 * @returns {{}} The child.
 */
function appendChildIfNotFound(node, name) {
    let child = findChild(node, name);
    if (!child) {
        child = { name, attributes: {}, children: [] };
        appendChild(node, child);
    }
    return child;
}
exports.appendChildIfNotFound = appendChildIfNotFound;
/**
 * Find a child with the given name.
 * @param {{}} node - The parent node.
 * @param {string} name - The name to find.
 * @returns {undefined|{}} The child if found.
 */
function findChild(node, name) {
    return _.find(node.children, { name });
}
exports.findChild = findChild;
/**
 * Get an attribute from a child node.
 * @param {{}} node - The parent node.
 * @param {string} name - The name of the child node.
 * @param {string} attribute - The name of the attribute.
 * @returns {undefined|*} The value of the attribute if found.
 */
function getChildAttribute(node, name, attribute) {
    const child = findChild(node, name);
    if (child)
        return child.attributes && child.attributes[attribute];
}
exports.getChildAttribute = getChildAttribute;
/**
 * Returns a value indicating whether the node has a child with the given name.
 * @param {{}} node - The parent node.
 * @param {string} name - The name of the child node.
 * @returns {boolean} True if found, false otherwise.
 */
function hasChild(node, name) {
    return _.some(node.children, { name });
}
exports.hasChild = hasChild;
/**
 * Insert the child after the specified node.
 * @param {{}} node - The parent node.
 * @param {{}} child - The child node.
 * @param {{}} after - The node to insert after.
 * @returns {undefined}
 */
function insertAfter(node, child, after) {
    if (!node.children)
        node.children = [];
    const index = node.children.indexOf(after);
    node.children.splice(index + 1, 0, child);
}
exports.insertAfter = insertAfter;
/**
 * Insert the child before the specified node.
 * @param {{}} node - The parent node.
 * @param {{}} child - The child node.
 * @param {{}} before - The node to insert before.
 * @returns {undefined}
 */
function insertBefore(node, child, before) {
    if (!node.children)
        node.children = [];
    const index = node.children.indexOf(before);
    node.children.splice(index, 0, child);
}
exports.insertBefore = insertBefore;
/**
 * Insert a child node in the correct order.
 * @param {{}} node - The parent node.
 * @param {{}} child - The child node.
 * @param {Array.<string>} nodeOrder - The order of the node names.
 * @returns {undefined}
 */
function insertInOrder(node, child, nodeOrder) {
    const childIndex = nodeOrder.indexOf(child.name);
    if (node.children && childIndex >= 0) {
        for (let i = childIndex + 1; i < nodeOrder.length; i++) {
            const sibling = findChild(node, nodeOrder[i]);
            if (sibling) {
                insertBefore(node, child, sibling);
                return;
            }
        }
    }
    appendChild(node, child);
}
exports.insertInOrder = insertInOrder;
/**
 * Check if the node is empty (no attributes and no children).
 * @param {{}} node - The node.
 * @returns {boolean} True if empty, false otherwise.
 */
function isEmpty(node) {
    return _.isEmpty(node.children) && _.isEmpty(node.attributes);
}
exports.isEmpty = isEmpty;
/**
 * Remove a child node.
 * @param {{}} node - The parent node.
 * @param {string|{}} child - The child node or name of node.
 * @returns {undefined}
 */
function removeChild(node, child) {
    if (!node.children)
        return;
    if (typeof child === 'string') {
        _.remove(node.children, { name: child });
    }
    else {
        const index = node.children.indexOf(child);
        if (index >= 0)
            node.children.splice(index, 1);
    }
}
exports.removeChild = removeChild;
/**
 * Set/unset the attributes on the node.
 * @param {{}} node - The node.
 * @param {{}} attributes - The attributes to set.
 * @returns {undefined}
 */
function setAttributes(node, attributes) {
    _.forOwn(attributes, (value, attribute) => {
        if (_.isNil(value)) {
            if (node.attributes)
                delete node.attributes[attribute];
        }
        else {
            if (!node.attributes)
                node.attributes = {};
            node.attributes[attribute] = value;
        }
    });
}
exports.setAttributes = setAttributes;
/**
 * Set attributes on a child node, creating the child if necessary.
 * @param {{}} node - The parent node.
 * @param {string} name - The name of the child node.
 * @param {{}} attributes - The attributes to set.
 * @returns {{}} The child.
 */
function setChildAttributes(node, name, attributes) {
    let child = findChild(node, name);
    _.forOwn(attributes, (value, attribute) => {
        if (_.isNil(value)) {
            if (child && child.attributes)
                delete child.attributes[attribute];
        }
        else {
            if (!child) {
                child = { name, attributes: {}, children: [] };
                appendChild(node, child);
            }
            if (!child.attributes)
                child.attributes = {};
            child.attributes[attribute] = value;
        }
    });
    return child;
}
exports.setChildAttributes = setChildAttributes;
/**
 * Remove the child node if empty.
 * @param {{}} node - The parent node.
 * @param {string|{}} child - The child or name of child node.
 * @returns {undefined}
 */
function removeChildIfEmpty(node, child) {
    if (typeof child === 'string')
        child = findChild(node, child);
    if (child && isEmpty(child))
        removeChild(node, child);
}
exports.removeChildIfEmpty = removeChildIfEmpty;
//# sourceMappingURL=xmlq.js.map