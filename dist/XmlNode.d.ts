export interface IXmlAttributes {
    [index: string]: string | number;
}
export declare type XmlChild = XmlNode | string | number;
export declare class XmlNode {
    name: string;
    children?: XmlChild[];
    attributes?: IXmlAttributes;
    constructor(name: string, attributes?: IXmlAttributes);
    setAttributes(attributes: IXmlAttributes): void;
    appendChild(child: XmlChild): void;
    findChildWithName(name: string): XmlNode | undefined;
    hasChild(name: string): boolean;
    removeChild(child: XmlNode): boolean;
    removeChildWithName(name: string): boolean;
    toString(includeDeclaration?: boolean): string;
}
//# sourceMappingURL=XmlNode.d.ts.map