import { INode } from './XmlParser';

const RELATIONSHIP_SCHEMA_PREFIX = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';

/**
 * A relationship collection.
 */
export class Relationships {
    private readonly node: INode;
    private nextId: number;

    /**
     * Creates a new instance of _Relationships.
     * @param node - The node.
     */
    public constructor(node?: INode) {
        this.node = node || {
            name: 'Relationships',
            attributes: {
                xmlns: 'http://schemas.openxmlformats.org/package/2006/relationships',
            },
        };

        this.nextId = 1;
        if (this.node.children) {
            this.node.children.forEach(child => {
                if (typeof child !== 'string'
                    && typeof child !== 'number'
                    && child.attributes
                    && typeof child.attributes.Id === 'string') {
                    const id = parseInt(child.attributes.Id.substr(3), 10);
                    if (id >= this.nextId) this.nextId = id + 1;
                }
            });
        }
    }

    /**
     * Add a new relationship.
     * @param type - The type of relationship.
     * @param target - The target of the relationship.
     * @param [targetMode] - The target mode of the relationship.
     * @returns The new relationship.
     */
    public add(type: string, target: string, targetMode?: string): INode {
        const node = {
            name: 'Relationship',
            attributes: {
                Id: `rId${this.nextId++}`,
                Type: `${RELATIONSHIP_SCHEMA_PREFIX}${type}`,
                Target: target,
            },
        };

        if (targetMode) {
            (node.attributes as any).TargetMode = targetMode;
        }

        if (!this.node.children) this.node.children = [];
        this.node.children.push(node);
        return node;
    }

    /**
     * Find a relationship by ID.
     * @param id - The relationship ID.
     * @returns The matching relationship or undefined if not found.
     */
    public findById(id: string): INode|undefined {
        return this.node.children && this.node.children.find(node => {
            return !!(typeof node !== 'string'
                && typeof node !== 'number'
                && node.attributes
                && node.attributes.Id === id);
        }) as INode|undefined;
    }

    /**
     * Find a relationship by type.
     * @param type - The type to search for.
     * @returns The matching relationship or undefined if not found.
     */
    public findByType(type: string): INode|undefined {
        return this.node.children && this.node.children.find(node => {
            return !!(typeof node !== 'string'
                && typeof node !== 'number'
                && node.attributes
                && node.attributes.Type === `${RELATIONSHIP_SCHEMA_PREFIX}${type}`);
        }) as INode|undefined;
    }

    /**
     * Convert the collection to an XML object.
     * @returns The XML or undefined if empty.
     */
    public toXml(): INode|undefined {
        if (!this.node.children || !this.node.children.length) return;
        return this.node;
    }
}

// tslint:disable
/*
xl/_rels/workbook.xml.rels

<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
    <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
    <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>
*/

