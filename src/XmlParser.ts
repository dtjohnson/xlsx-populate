/**
 * @module xlsx-populate
 */

import { SAXParser } from 'sax';

// Regex to check if string is all whitespace.
const allWhitespaceRegex = /^\s+$/;

export interface INode {
    name: string;
    attributes: {
        [index: string]: string|number;
    };
    children: (INode|string|number)[];
}

/**
 * XML parser.
 * @ignore
 */
export class XmlParser {
    /**
     * Parse the XML text into a JSON object.
     * @param xmlText - The XML text.
     * @returns The JSON object.
     */
    public parseAsync(xmlText: string): Promise<INode> {
        return new Promise((resolve, reject) => {
            // Create the SAX parser.
            const parser = new SAXParser(true);

            // Parsed is the full parsed object. Current is the current node being parsed. Stack is the current stack of
            // nodes leading to the current one.
            let parsed: INode, current: INode;
            const stack: INode[] = [];

            // On error: Reject the promise.
            parser.onerror = reject;

            // On text nodes: If it is all whitespace, do nothing. Otherwise, try to convert to a number and add as a child.
            parser.ontext = (text: string) => {
                if (allWhitespaceRegex.test(text)) {
                    if (current && current.attributes['xml:space'] === 'preserve') {
                        current.children.push(text);
                    }
                } else {
                    current.children.push(this.covertToNumberIfNumber(text));
                }
            };

            // On open tag start: Create a child element. If this is the root element, set it as parsed. Otherwise, add
            // it as a child to the current node.
            (parser as any).onopentagstart = (node: { name: string }) => {
                const child = { name: node.name, attributes: {}, children: [] };
                if (current) {
                    current.children.push(child);
                } else {
                    parsed = child;
                }

                stack.push(child);
                current = child;
            };

            // On close tag: Pop the stack.
            parser.onclosetag = (_tagName: string) => {
                stack.pop();
                current = stack[stack.length - 1];
            };

            // On attribute: Try to convert the value to a number and add to the current node.
            parser.onattribute = (attribute: { name: string; value: string }) => {
                current.attributes[attribute.name] = this.covertToNumberIfNumber(attribute.value);
            };

            // On end: Resolve the promise.
            parser.onend = () => resolve(parsed);

            // Start parsing the text.
            parser.write(xmlText).close();
        });
    }

    /**
     * Convert the string to a number if it looks like one.
     * @param str - The string to convert.
     * @returns The number if converted or the string if not.
     */
    private covertToNumberIfNumber(str: string): number|string {
        const num = Number(str);
        return num.toString() === str ? num : str;
    }
}
