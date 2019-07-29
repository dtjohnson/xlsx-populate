import { INode } from './XmlParser';
import * as xmlq from './xmlq';

/**
 * A content type collection.
 * @ignore
 */
export class ContentTypes {
    /**
     * Creates a new instance of ContentTypes
     * @param node - The node.
     */
    public constructor(private readonly node: INode) {}

    /**
     * Add a new content type.
     * @param partName - The part name.
     * @param contentType - The content type.
     * @returns The new content type.
     */
    public add(partName: string, contentType: string): INode {
        const node = {
            name: 'Override',
            attributes: {
                PartName: partName,
                ContentType: contentType,
            },
        };

        xmlq.appendChild(this.node, node);

        return node;
    }

    /**
     * Find a content type by part name.
     * @param partName - The part name.
     * @returns The matching content type or undefined if not found.
     */
    public findByPartName(partName: string): INode|undefined {
        if (!this.node.children) return;
        for (const node of this.node.children) {
            if (typeof node !== 'string' && typeof node !== 'number' && node.attributes && node.attributes.PartName === partName) {
                return node;
            }
        }
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
[Content_Types].xml

<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"/>
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
    <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
    <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
    <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
    <Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>
    <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
    <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
*/
