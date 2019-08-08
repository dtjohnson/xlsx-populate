import { COLORS } from './colors';
import { Color } from './types';
import { INode } from './XmlParser';
import * as xmlq from './xmlq';

export function getColor(node: INode, name: string): Color|undefined {
    const child = xmlq.findChild(node, name);
    if (!child || !child.attributes) return;

    let color: Color;
    if (child.attributes.hasOwnProperty('rgb')) {
        color = {
            rgb : String(child.attributes.rgb),
        };
    } else if (child.attributes.hasOwnProperty('theme')) {
        color = {
            theme : Number(child.attributes.theme),
        };
    } else if (child.attributes.hasOwnProperty('indexed')) {
        color = {
            rgb : COLORS[Number(child.attributes.indexed)],
        };
    } else {
        return;
    }

    if (child.attributes.hasOwnProperty('tint')) color.tint = Number(child.attributes.tint);

    return color;
}

export function setColor(node: INode, name: string, color: Color|undefined): void {
    const attributes = {
        rgb: undefined as string|undefined,
        indexed: undefined,
        theme: undefined as number|undefined,
        tint: undefined as number|undefined,
    };

    if (color && 'rgb' in color) attributes.rgb = color.rgb.toUpperCase();
    if (color && 'theme' in color) attributes.theme = color.theme;
    if (color && 'tint' in color) attributes.tint = color.tint;

    xmlq.setChildAttributes(node, name, attributes);
    xmlq.removeChildIfEmpty(node, 'color');
}
