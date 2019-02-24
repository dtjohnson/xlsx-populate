/**
 * Append a child to the node.
 * @param {{}} node - The parent node.
 * @param {{}} child - The child node.
 * @returns {undefined}
 */
export declare function appendChild(node: any, child: any): void;
/**
 * Append a child if one with the given name is not found.
 * @param {{}} node - The parent node.
 * @param {string} name - The child node name.
 * @returns {{}} The child.
 */
export declare function appendChildIfNotFound(node: any, name: any): any;
/**
 * Find a child with the given name.
 * @param {{}} node - The parent node.
 * @param {string} name - The name to find.
 * @returns {undefined|{}} The child if found.
 */
export declare function findChild(node: any, name: any): any;
/**
 * Get an attribute from a child node.
 * @param {{}} node - The parent node.
 * @param {string} name - The name of the child node.
 * @param {string} attribute - The name of the attribute.
 * @returns {undefined|*} The value of the attribute if found.
 */
export declare function getChildAttribute(node: any, name: any, attribute: any): any;
/**
 * Returns a value indicating whether the node has a child with the given name.
 * @param {{}} node - The parent node.
 * @param {string} name - The name of the child node.
 * @returns {boolean} True if found, false otherwise.
 */
export declare function hasChild(node: any, name: any): any;
/**
 * Insert the child after the specified node.
 * @param {{}} node - The parent node.
 * @param {{}} child - The child node.
 * @param {{}} after - The node to insert after.
 * @returns {undefined}
 */
export declare function insertAfter(node: any, child: any, after: any): void;
/**
 * Insert the child before the specified node.
 * @param {{}} node - The parent node.
 * @param {{}} child - The child node.
 * @param {{}} before - The node to insert before.
 * @returns {undefined}
 */
export declare function insertBefore(node: any, child: any, before: any): void;
/**
 * Insert a child node in the correct order.
 * @param {{}} node - The parent node.
 * @param {{}} child - The child node.
 * @param {Array.<string>} nodeOrder - The order of the node names.
 * @returns {undefined}
 */
export declare function insertInOrder(node: any, child: any, nodeOrder: any): void;
/**
 * Check if the node is empty (no attributes and no children).
 * @param {{}} node - The node.
 * @returns {boolean} True if empty, false otherwise.
 */
export declare function isEmpty(node: any): any;
/**
 * Remove a child node.
 * @param {{}} node - The parent node.
 * @param {string|{}} child - The child node or name of node.
 * @returns {undefined}
 */
export declare function removeChild(node: any, child: any): void;
/**
 * Set/unset the attributes on the node.
 * @param {{}} node - The node.
 * @param {{}} attributes - The attributes to set.
 * @returns {undefined}
 */
export declare function setAttributes(node: any, attributes: any): void;
/**
 * Set attributes on a child node, creating the child if necessary.
 * @param {{}} node - The parent node.
 * @param {string} name - The name of the child node.
 * @param {{}} attributes - The attributes to set.
 * @returns {{}} The child.
 */
export declare function setChildAttributes(node: any, name: any, attributes: any): any;
/**
 * Remove the child node if empty.
 * @param {{}} node - The parent node.
 * @param {string|{}} child - The child or name of child node.
 * @returns {undefined}
 */
export declare function removeChildIfEmpty(node: any, child: any): void;
//# sourceMappingURL=xmlq.d.ts.map