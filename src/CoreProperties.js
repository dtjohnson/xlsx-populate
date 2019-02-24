"use strict";

const allowedProperties = {
    title: "dc:title",
    subject: "dc:subject",
    author: "dc:creator",
    creator: "dc:creator",
    description: "dc:description",
    keywords: "cp:keywords",
    category: "cp:category"
};

/**
 * Core properties
 * @ignore
 */
class CoreProperties {
    constructor(node) {
        this._node = node;
        this._properties = {};
    }

    /**
     * Sets a specific property.
     * @param {string} name - The name of the property.
     * @param {*} value - The value of the property.
     * @returns {CoreProperties} CoreProperties.
     */
    set(name, value) {
        const key = name.toLowerCase();

        if (typeof allowedProperties[key] === "undefined") {
            throw new Error(`Unknown property name: "${name}"`);
        }

        this._properties[key] = value;

        return this;
    }

    /**
     * Get a specific property.
     * @param {string} name - The name of the property.
     * @returns {*} The property value.
     */
    get(name) {
        const key = name.toLowerCase();

        if (typeof allowedProperties[key] === "undefined") {
            throw new Error(`Unknown property name: "${name}"`);
        }

        return this._properties[key];
    }

    /**
     * Convert the collection to an XML object.
     * @returns {{}} The XML.
     */
    toXml() {
        for (const key in this._properties) {
            if (!this._properties.hasOwnProperty(key)) continue;
            this._node.children.push({
                name: allowedProperties[key],
                children: [this._properties[key]]
            });
        }

        return this._node;
    }
}

module.exports = CoreProperties;

/*
docProps/core.xml

<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:title>Title</dc:title>
<dc:subject>Subject</dc:subject>
<dc:creator>Creator</dc:creator>
<cp:keywords>Keywords</cp:keywords>
<dc:description>Description</dc:description>
<cp:category>Category</cp:category>
</cp:coreProperties>
 */
