import _ from 'lodash';
import { INode, NodeAttributes } from './XmlParser';
import * as xmlq from './xmlq';

const propertyNameMap = {
    title: 'dc:title',
    subject: 'dc:subject',
    author: 'dc:creator',
    keywords: 'cp:keywords',
    comments: 'dc:description',
    lastModifiedBy: 'cp:lastModifiedBy',
    created: 'dcterms:created',
    modified: 'dcterms:modified',
    category: 'cp:category',
};

type PropertyName = keyof typeof propertyNameMap;

/**
 * Core workbook properties
 */
export class CoreProperties {
    public constructor(
        private readonly node: INode,
    ) {}

    public get title(): string|undefined {
        return this.getProperty('title');
    }

    public set title(value: string|undefined) {
        this.setProperty('title', value);
    }

    public get subject(): string|undefined {
        return this.getProperty('subject');
    }

    public set subject(value: string|undefined) {
        this.setProperty('subject', value);
    }

    public get author(): string|undefined {
        return this.getProperty('author');
    }

    public set author(value: string|undefined) {
        this.setProperty('author', value);
    }

    public get keywords(): string|undefined {
        return this.getProperty('keywords');
    }

    public set keywords(value: string|undefined) {
        this.setProperty('keywords', value);
    }

    public get comments(): string|undefined {
        return this.getProperty('comments');
    }

    public set comments(value: string|undefined) {
        this.setProperty('comments', value);
    }

    public get lastModifiedBy(): string|undefined {
        return this.getProperty('lastModifiedBy');
    }

    public set lastModifiedBy(value: string|undefined) {
        this.setProperty('lastModifiedBy', value);
    }

    public get created(): Date|undefined {
        return this.getDateProperty('created');
    }

    public set created(value: Date|undefined) {
        this.setDateProperty('created', value);
    }

    public get modified(): Date|undefined {
        return this.getDateProperty('modified');
    }

    public set modified(value: Date|undefined) {
        this.setDateProperty('modified', value);
    }

    public get category(): string|undefined {
        return this.getProperty('category');
    }

    public set category(value: string|undefined) {
        this.setProperty('category', value);
    }

    /**
     * Convert the collection to an XML object.
     * @returns The XML.
     */
    public toXml(): INode {
        return this.node;
    }

    private setDateProperty(name: PropertyName, value: Date|undefined): void {
        const strValue = value && `${value.toISOString().split('.')[0]}Z`;
        this.setProperty(name, strValue, { 'xsi:type': 'dcterms:W3CDTF' });
    }

    private getDateProperty(name: PropertyName): Date|undefined {
        const strValue = this.getProperty(name);
        return _.isNil(strValue) ? undefined : new Date(strValue);
    }

    private setProperty(name: PropertyName, value: string|undefined, attributes?: NodeAttributes): void {
        const nodeName = propertyNameMap[name];

        if (_.isNil(value)) {
            xmlq.removeChild(this.node, nodeName);
        } else {
            const propertyNode = xmlq.appendChildIfNotFound(this.node, nodeName);
            propertyNode.children = [ value ];
            if (attributes) {
                xmlq.setAttributes(propertyNode, attributes);
            }
        }
    }

    private getProperty(name: PropertyName): string|undefined {
        const nodeName = propertyNameMap[name];
        const propertyNode = xmlq.findChild(this.node, nodeName);
        if (!propertyNode || !propertyNode.children || !propertyNode.children[0]) return;
        return String(propertyNode.children[0]);
    }
}

// tslint:disable
/*
docProps/core.xml

<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>TITLE</dc:title>
  <dc:subject>SUBJECT</dc:subject>
  <dc:creator>AUTHOR</dc:creator>
  <cp:keywords>KEYWORDS</cp:keywords>
  <dc:description>COMMENTS</dc:description>
  <cp:lastModifiedBy>LAST_MODIFIED_BY</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">2019-08-09T11:44:11Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2019-08-09T11:44:11Z</dcterms:modified>
  <cp:category>CATEGORY</cp:category>
</cp:coreProperties>
 */
