interface INode {
    name: string;
    attributes: {
        [index: string]: string | number;
    };
    children: (INode | string | number)[];
}
/**
 * XML parser.
 * @private
 */
export declare class XmlParser {
    /**
     * Parse the XML text into a JSON object.
     * @param {string} xmlText - The XML text.
     * @returns {{}} The JSON object.
     */
    parseAsync(xmlText: string): Promise<INode>;
    /**
     * Convert the string to a number if it looks like one.
     * @param {string} str - The string to convert.
     * @returns {string|number} The number if converted or the string if not.
     * @private
     */
    _covertToNumberIfNumber(str: string): number | string;
}
export {};
//# sourceMappingURL=XmlParser.d.ts.map