"use strict";

const _ = require("lodash");
const xmlq = require("./xmlq");
const ArgHandler = require("./ArgHandler");

/**
 * App properties
 * @ignore
 */
class AppProperties {
    /**
     * Creates a new instance of AppProperties
     * @param {{}} node - The node.
     */
    constructor(node) {
        this._node = node;
    }

    isSecure(value) {
        return new ArgHandler("Range.formula")
            .case(() => {
                const docSecurityNode = xmlq.findChild(this._node, "DocSecurity");
                if (!docSecurityNode) return false;
                return docSecurityNode.children[0] === 1;
            })
            .case('boolean', value => {
                const docSecurityNode = xmlq.appendChildIfNotFound(this._node, "DocSecurity");
                docSecurityNode.children = [value ? 1 : 0];
                return this;
            })
            .handle(arguments);
    }

    /**
     * Convert the collection to an XML object.
     * @returns {{}} The XML.
     */
    toXml() {
        return this._node;
    }
}

module.exports = AppProperties;

/*
docProps/app.xml

<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
    <Application>Microsoft Excel</Application>
<DocSecurity>1</DocSecurity>
<ScaleCrop>false</ScaleCrop>
<HeadingPairs>
<vt:vector size="2" baseType="variant">
    <vt:variant>
<vt:lpstr>Worksheets</vt:lpstr>
</vt:variant>
<vt:variant>
<vt:i4>1</vt:i4>
</vt:variant>
</vt:vector>
</HeadingPairs>
<TitlesOfParts>
<vt:vector size="1" baseType="lpstr">
    <vt:lpstr>Sheet1</vt:lpstr>
</vt:vector>
</TitlesOfParts>
<Company/>
<LinksUpToDate>false</LinksUpToDate>
<SharedDoc>false</SharedDoc>
<HyperlinksChanged>false</HyperlinksChanged>
<AppVersion>16.0300</AppVersion>
</Properties>
 */
