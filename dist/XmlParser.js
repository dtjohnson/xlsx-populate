"use strict";
/**
 * @module xlsx-populate
 */
Object.defineProperty(exports, "__esModule", { value: true });
const sax_1 = require("sax");
// Regex to check if string is all whitespace.
const allWhitespaceRegex = /^\s+$/;
/**
 * XML parser.
 * @ignore
 */
class XmlParser {
    /**
     * Parse the XML text into a JSON object.
     * @param xmlText - The XML text.
     * @returns The JSON object.
     */
    parseAsync(xmlText) {
        return new Promise((resolve, reject) => {
            // Create the SAX parser.
            const parser = new sax_1.SAXParser(true);
            // Parsed is the full parsed object. Current is the current node being parsed. Stack is the current stack of
            // nodes leading to the current one.
            let parsed, current;
            const stack = [];
            // On error: Reject the promise.
            parser.onerror = reject;
            // On text nodes: If it is all whitespace, do nothing. Otherwise, try to convert to a number and add as a child.
            parser.ontext = (text) => {
                if (allWhitespaceRegex.test(text)) {
                    if (current && current.attributes && current.attributes['xml:space'] === 'preserve') {
                        if (!current.children)
                            current.children = [];
                        current.children.push(text);
                    }
                }
                else {
                    if (!current.children)
                        current.children = [];
                    current.children.push(this.covertToNumberIfNumber(text));
                }
            };
            // On open tag start: Create a child element. If this is the root element, set it as parsed. Otherwise, add
            // it as a child to the current node.
            parser.onopentagstart = (node) => {
                const child = { name: node.name };
                if (current) {
                    if (!current.children)
                        current.children = [];
                    current.children.push(child);
                }
                else {
                    parsed = child;
                }
                stack.push(child);
                current = child;
            };
            // On close tag: Pop the stack.
            parser.onclosetag = (_tagName) => {
                stack.pop();
                current = stack[stack.length - 1];
            };
            // On attribute: Try to convert the value to a number and add to the current node.
            parser.onattribute = (attribute) => {
                if (!current.attributes)
                    current.attributes = {};
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
    covertToNumberIfNumber(str) {
        const num = Number(str);
        return num.toString() === str ? num : str;
    }
}
exports.XmlParser = XmlParser;
//# sourceMappingURL=XmlParser.js.map