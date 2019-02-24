const _ = require("lodash");

/**
 * Append a child to the node.
 * @param {{}} node - The parent node.
 * @param {{}} child - The child node.
 * @returns {undefined}
 */
export function appendChild(node: any, child: any) {
    if (!node.children) node.children = [];
    node.children.push(child);
}

/**
 * Append a child if one with the given name is not found.
 * @param {{}} node - The parent node.
 * @param {string} name - The child node name.
 * @returns {{}} The child.
 */
export function appendChildIfNotFound(node: any, name: any) {
    let child = findChild(node, name);
    if (!child) {
        child = { name, attributes: {}, children: [] };
        appendChild(node, child);
    }

    return child;
}

/**
 * Find a child with the given name.
 * @param {{}} node - The parent node.
 * @param {string} name - The name to find.
 * @returns {undefined|{}} The child if found.
 */
export function findChild(node: any, name: any) {
    return _.find(node.children, { name });
}

/**
 * Get an attribute from a child node.
 * @param {{}} node - The parent node.
 * @param {string} name - The name of the child node.
 * @param {string} attribute - The name of the attribute.
 * @returns {undefined|*} The value of the attribute if found.
 */
export function getChildAttribute(node: any, name: any, attribute: any) {
    const child = findChild(node, name);
    if (child) return child.attributes && child.attributes[attribute];
}

/**
 * Returns a value indicating whether the node has a child with the given name.
 * @param {{}} node - The parent node.
 * @param {string} name - The name of the child node.
 * @returns {boolean} True if found, false otherwise.
 */
export function hasChild(node: any, name: any) {
    return _.some(node.children, { name });
}

/**
 * Insert the child after the specified node.
 * @param {{}} node - The parent node.
 * @param {{}} child - The child node.
 * @param {{}} after - The node to insert after.
 * @returns {undefined}
 */
export function insertAfter(node: any, child: any, after: any) {
    if (!node.children) node.children = [];
    const index = node.children.indexOf(after);
    node.children.splice(index + 1, 0, child);
}

/**
 * Insert the child before the specified node.
 * @param {{}} node - The parent node.
 * @param {{}} child - The child node.
 * @param {{}} before - The node to insert before.
 * @returns {undefined}
 */
export function insertBefore(node: any, child: any, before: any) {
    if (!node.children) node.children = [];
    const index = node.children.indexOf(before);
    node.children.splice(index, 0, child);
}

/**
 * Insert a child node in the correct order.
 * @param {{}} node - The parent node.
 * @param {{}} child - The child node.
 * @param {Array.<string>} nodeOrder - The order of the node names.
 * @returns {undefined}
 */
export function insertInOrder(node: any, child: any, nodeOrder: any) {
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

/**
 * Check if the node is empty (no attributes and no children).
 * @param {{}} node - The node.
 * @returns {boolean} True if empty, false otherwise.
 */
export function isEmpty(node: any) {
    return _.isEmpty(node.children) && _.isEmpty(node.attributes);
}

/**
 * Remove a child node.
 * @param {{}} node - The parent node.
 * @param {string|{}} child - The child node or name of node.
 * @returns {undefined}
 */
export function removeChild(node: any, child: any) {
    if (!node.children) return;
    if (typeof child === 'string') {
        _.remove(node.children, { name: child });
    } else {
        const index = node.children.indexOf(child);
        if (index >= 0) node.children.splice(index, 1);
    }
}

/**
 * Set/unset the attributes on the node.
 * @param {{}} node - The node.
 * @param {{}} attributes - The attributes to set.
 * @returns {undefined}
 */
export function setAttributes(node: any, attributes: any) {
    _.forOwn(attributes, (value: any, attribute: any) => {
        if (_.isNil(value)) {
            if (node.attributes) delete node.attributes[attribute];
        } else {
            if (!node.attributes) node.attributes = {};
            node.attributes[attribute] = value;
        }
    });
}

/**
 * Set attributes on a child node, creating the child if necessary.
 * @param {{}} node - The parent node.
 * @param {string} name - The name of the child node.
 * @param {{}} attributes - The attributes to set.
 * @returns {{}} The child.
 */
export function setChildAttributes(node: any, name: any, attributes: any) {
    let child = findChild(node, name);
    _.forOwn(attributes, (value: any, attribute: any) => {
        if (_.isNil(value)) {
            if (child && child.attributes) delete child.attributes[attribute];
        } else {
            if (!child) {
                child = { name, attributes: {}, children: [] };
                appendChild(node, child);
            }

            if (!child.attributes) child.attributes = {};
            child.attributes[attribute] = value;
        }
    });

    return child;
}

/**
 * Remove the child node if empty.
 * @param {{}} node - The parent node.
 * @param {string|{}} child - The child or name of child node.
 * @returns {undefined}
 */
export function removeChildIfEmpty(node: any, child: any) {
    if (typeof child === 'string') child = findChild(node, child);
    if (child && isEmpty(child)) removeChild(node, child);
}
