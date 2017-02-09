"use strict";

const builder = require('xmlbuilder');
const _ = require("lodash");

class XmlBuilder {
    build(node) {
        // const xml = builder.create(node.name, { encoding: "UTF-8", standalone: true })

        const xml = this._build(node);
        return xml.end({ pretty: true });
    }

    _build(node, xml) {
        if (typeof node === "string") {
            xml.text(node);
        } else {
            if (xml) {
                xml = xml.element(node.name);
            } else {
                xml = builder.create(node.name, { encoding: "UTF-8", standalone: true });
            }

            _.forOwn(node.attributes, (value, name) => {
                xml.attribute(name, value);
            });

            _.forEach(node.children, child => this._build(child, xml));
        }

        return xml;
    }
};

module.exports = XmlBuilder;
