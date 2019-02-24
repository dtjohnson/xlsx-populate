// TODO: This is a reworked API for XML nodes.

const XML_DECLARATION = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`;

export interface XmlAttributes {
    [index:string]: string|number
}

export type XmlChild = XmlNode|string|number;

/**
 * Escape a string for use in XML by replacing &, ", ', <, and >.
 * @param {*} value - The value to escape.
 * @param {boolean} [isAttribute] - A flag indicating if this is an attribute.
 * @returns {string} The escaped string.
 * @private
 */
function escapeString(value: number|string, isAttribute: boolean) {
    value = value.toString()
        .replace(/&/g, "&amp;") // Escape '&' first as the other escapes add them.
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;");

    if (isAttribute) {
        value = value.replace(/"/g, "&quot;");
    }

    return value;
}

export class XmlNode {
    public name: string;

    public children?: XmlChild[];

    public attributes?: XmlAttributes;

    public constructor(name: string, attributes?: XmlAttributes) {
        this.name = name;
        if (attributes) this.setAttributes(attributes);
    }

    public setAttributes(attributes: XmlAttributes) {
        if (!this.attributes) this.attributes = {};
        for (let key in attributes) {
            if (attributes.hasOwnProperty(key)) {
                this.attributes[key] = attributes[key];
            }
        }
    }

    public appendChild(child: XmlChild): void {
        if (!this.children) this.children = [];
        this.children.push(child);
    }

    public findChildWithName(name: string): XmlNode|undefined {
        if (!this.children) return;
        for (let i = 0; i < this.children.length; i++) {
            const child = this.children[i];
            if (child instanceof XmlNode && child.name === name) return child;
        }
    }

    public hasChild(name: string): boolean {
        return !!this.findChildWithName(name);
    }

    public removeChild(child: XmlNode): boolean {
        if (!this.children) return false;
        const index = this.children.indexOf(child);
        if (index < 0) return false;
        this.children.splice(index, 1);
        return true;
    }

    public removeChildWithName(name: string): boolean {
        const child = this.findChildWithName(name);
        if (!child) return false;
        return this.removeChild(child);
    }

    public toString(includeDeclaration: boolean = false): string {
        let str = `<${this.name}`;

        if (this.attributes) {
            for (let key in this.attributes) {
                if (this.attributes.hasOwnProperty(key)) {
                    str += ` ${key}="${escapeString(this.attributes[key], true)}"`;
                }
            }
        }

        if (this.children && this.children.length) {
            str += ">";

            // Recursively add any children.
            this.children.forEach(child => {
                str += (child instanceof XmlNode) ? child : escapeString(child, false);
            });

            // Close the tag.
            str += `</${this.name}>`;
        } else {
            // Self-close the tag if no children.
            str += "/>";
        }

        if (includeDeclaration) {
            str = XML_DECLARATION + str;
        }

        return str;
    }
}
