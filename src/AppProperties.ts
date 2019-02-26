/**
 * @module xlsx-populate
 */

import { OverloadHandler } from './OverloadHandler';
import { INode } from './XmlParser';
import * as xmlq from './xmlq';

/**
 * App properties
 * @ignore
 */
export class AppProperties {
    /**
     * Creates a new instance of AppProperties
     * @param node - The node.
     */
    public constructor(private node: INode) {}

    /**
     * Get a value indicating whether the workbook is secure or not.
     */
    public isSecure(): boolean;
    /**
     * Set a value indicating whether the workbook is secure or not.
     * @param value - The value to set.
     */
    public isSecure(value: boolean): this;
    public isSecure(...args: any[]): any {
        return new OverloadHandler('AppProperties.isSecure')
            .case<boolean>(() => {
                const docSecurityNode = xmlq.findChild(this.node, 'DocSecurity');
                if (!docSecurityNode) return false;
                return !!(docSecurityNode.children && docSecurityNode.children.length && docSecurityNode.children[0] === 1);
            })
            .case<boolean, this>('boolean', value => {
                const docSecurityNode = xmlq.appendChildIfNotFound(this.node, 'DocSecurity');
                docSecurityNode.children = [ value ? 1 : 0 ];
                return this;
            })
            .handle(args);
    }

    /**
     * Convert the collection to an XML object.
     * @returns The XML.
     */
    public toXml(): INode {
        return this.node;
    }
}

// tslint:disable
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
