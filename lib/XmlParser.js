"use strict";

const Promise = require("bluebird");
const sax = require("sax");

const allWhitespaceRegex = /^\s+$/;

class XmlParser {
    parseAsync(xmlText) {
        return new Promise((resolve, reject) => {
            const parser = sax.parser(true);
            let parsed, current;
            const stack = [];

            parser.onerror = reject;

            parser.ontext = text => {
                if (allWhitespaceRegex.test(text)) return;
                current.children.push(text);
            };

            parser.onopentagstart = node => {
                const child = { name: node.name, attributes: {}, children: [] };
                if (current) {
                    current.children.push(child);
                } else {
                    parsed = child;
                }

                stack.push(child);
                current = child;
            };

            parser.onclosetag = node => {
                stack.pop();
                current = stack[stack.length - 1];
            };

            parser.onattribute = attribute => {
                current.attributes[attribute.name] = attribute.value;
            };

            parser.onend = () => resolve(parsed);

            parser.write(xmlText).close();
        });
    }
}

module.exports = XmlParser;
