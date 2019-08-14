import _ from 'lodash';
import { Borders } from './Borders';
import { StyleSheet } from './StyleSheet';
import { Color, Fill, FillPattern, FontGenericFamily, FontScheme, GradientFill, GradientType, HorizontalAlignment,
    TextDirection, VerticalAlignment } from './types';
import { getColor, setColor } from './utils';
import { INode } from './XmlParser';
import * as xmlq from './xmlq';

const fontGenericFamilyNames: FontGenericFamily[] = [ 'serif', 'sans-serif', 'monospace' ];
const fontGenericFamilyIds = [ 1, 2, 3 ];

/**
 * A style.
 */
export class Style {
    private readonly _borders: Borders;

    /**
     * Creates a new instance of _Style.
     * @param styleSheet - The styleSheet.
     * @param id - The style ID.
     * @param xfNode - The xf node.
     * @param fontNode - The font node.
     * @param fillNode - The fill node.
     * @param borderNode - The border node.
     */
    public constructor(
        private readonly styleSheet: StyleSheet,
        private readonly _id: number,
        private readonly xfNode: INode,
        private readonly fontNode: INode,
        private readonly fillNode: INode,
        borderNode: INode,
    ) {
        this._borders = new Borders(borderNode);
    }

    /**
     * Gets the style ID.
     * @returns The ID.
     */
    public get id(): number {
        return this._id;
    }

    /**
     * Gets/sets whether the font is bold.
     */
    public get bold(): boolean {
        return xmlq.hasChild(this.fontNode, 'b');
    }

    public set bold(bold: boolean) {
        if (bold) xmlq.appendChildIfNotFound(this.fontNode, 'b');
        else xmlq.removeChild(this.fontNode, 'b');
    }

    /**
     * Gets/sets whether the font is italicized.
     */
    public get italic(): boolean {
        return xmlq.hasChild(this.fontNode, 'i');
    }

    public set italic(italic: boolean) {
        if (italic) xmlq.appendChildIfNotFound(this.fontNode, 'i');
        else xmlq.removeChild(this.fontNode, 'i');
    }

    /**
     * Gets/sets the type of underline.
     */
    public get underline(): 'double'|boolean {
        const uNode = xmlq.findChild(this.fontNode, 'u');
        if (!uNode) return false;
        const val = uNode.attributes && uNode.attributes.val;
        return (val === 'double') ? val : true;
    }

    public set underline(underline: 'double'|boolean) {
        if (underline) {
            const uNode = xmlq.appendChildIfNotFound(this.fontNode, 'u');
            const val = typeof underline === 'string' ? underline : undefined;
            xmlq.setAttributes(uNode, { val });
        } else {
            xmlq.removeChild(this.fontNode, 'u');
        }
    }

    /**
     * Gets/sets whether the font is struck through with a horizontal line.
     */
    public get strikethrough(): boolean {
        return xmlq.hasChild(this.fontNode, 'strike');
    }

    public set strikethrough(strikethrough: boolean) {
        if (strikethrough) xmlq.appendChildIfNotFound(this.fontNode, 'strike');
        else xmlq.removeChild(this.fontNode, 'strike');
    }

    private get fontVerticalAlignment(): string|undefined {
        return xmlq.getChildAttributeAsString(this.fontNode, 'vertAlign', 'val');
    }

    private set fontVerticalAlignment(alignment: string|undefined) {
        xmlq.setChildAttributes(this.fontNode, 'vertAlign', { val: alignment });
        xmlq.removeChildIfEmpty(this.fontNode, 'vertAlign');
    }

    /**
     * Gets/sets whether the font is subscript. Cannot be combined with superscript.
     */
    public get subscript(): boolean {
        return this.fontVerticalAlignment === 'subscript';
    }

    public set subscript(subscript: boolean) {
        this.fontVerticalAlignment = subscript ? 'subscript' : undefined;
    }

    /**
     * Gets/sets whether the font is superscript. Cannot be combined with subscript.
     */
    public get superscript(): boolean {
        return this.fontVerticalAlignment === 'superscript';
    }

    public set superscript(superscript: boolean) {
        this.fontVerticalAlignment = superscript ? 'superscript' : undefined;
    }

    /**
     * Gets/sets the font size in points. Must be greater than 0.
     */
    public get fontSize(): number|undefined {
        return xmlq.getChildAttributeAsNumber(this.fontNode, 'sz', 'val');
    }

    public set fontSize(size: number|undefined) {
        xmlq.setChildAttributes(this.fontNode, 'sz', { val: size });
        xmlq.removeChildIfEmpty(this.fontNode, 'sz');
    }

    /**
     * Gets/sets the name of the font family.
     */
    public get fontFamily(): string|undefined {
        return xmlq.getChildAttributeAsString(this.fontNode, 'name', 'val');
    }

    public set fontFamily(family: string|undefined) {
        xmlq.setChildAttributes(this.fontNode, 'name', { val: family });
        xmlq.removeChildIfEmpty(this.fontNode, 'name');
    }

    public get fontGenericFamily(): FontGenericFamily|undefined {
        const id = xmlq.getChildAttributeAsNumber(this.fontNode, 'family', 'val');
        if (_.isNil(id)) return;
        const idx = fontGenericFamilyIds.indexOf(id);
        return idx >= 0 ? fontGenericFamilyNames[idx] : undefined;
    }

    public set fontGenericFamily(genericFamily: FontGenericFamily|undefined) {
        const idx = _.isNil(genericFamily) ? undefined : fontGenericFamilyNames.indexOf(genericFamily);
        const id = idx !== undefined ? fontGenericFamilyIds[idx] : undefined;
        xmlq.setChildAttributes(this.fontNode, 'family', { val: id });
        xmlq.removeChildIfEmpty(this.fontNode, 'family');
    }

    /**
     * Gets/sets the font color.
     */
    public get fontColor(): Color|undefined {
        return getColor(this.fontNode, 'color');
    }

    public set fontColor(color: Color|undefined) {
        setColor(this.fontNode, 'color', color);
    }

    /**
     * Gets/sets the font scheme.
     */
    public get fontScheme(): FontScheme|undefined {
        return xmlq.getChildAttributeAsString(this.fontNode, 'scheme', 'val') as FontScheme|undefined;
    }

    public set fontScheme(scheme: FontScheme|undefined) {
        xmlq.setChildAttributes(this.fontNode, 'scheme', { val: scheme });
        xmlq.removeChildIfEmpty(this.fontNode, 'scheme');
    }

    /**
     * Gets/sets the horizontal alignment.
     */
    public get horizontalAlignment(): HorizontalAlignment|undefined {
        return xmlq.getChildAttributeAsString(this.xfNode, 'alignment', 'horizontal') as HorizontalAlignment|undefined;
    }

    public set horizontalAlignment(alignment: HorizontalAlignment|undefined) {
        xmlq.setChildAttributes(this.xfNode, 'alignment', { horizontal: alignment });
        xmlq.removeChildIfEmpty(this.xfNode, 'alignment');
    }

    /**
     * Gets/sets whether the last line of the cell is justified. Also known as 'Justified Distributed'. Only applies
     * when horizontalAlignment === 'distributed'. (This is typical for East Asian alignments but not typical in other
     * contexts.)
     */
    public get justifyLastLine(): boolean {
        return xmlq.getChildAttribute(this.xfNode, 'alignment', 'justifyLastLine') === 1;
    }

    public set justifyLastLine(justifyLastLine: boolean) {
        xmlq.setChildAttributes(this.xfNode, 'alignment', { justifyLastLine: justifyLastLine ? 1 : undefined });
        xmlq.removeChildIfEmpty(this.xfNode, 'alignment');
    }

    /**
     * Gets/sets the number of indents. Must be greater than or equal to 0.
     */
    public get indent(): number|undefined {
        return xmlq.getChildAttributeAsNumber(this.xfNode, 'alignment', 'indent');
    }

    public set indent(indent: number|undefined) {
        xmlq.setChildAttributes(this.xfNode, 'alignment', { indent });
        xmlq.removeChildIfEmpty(this.xfNode, 'alignment');
    }

    /**
     * Gets/sets the vertical alignment.
     */
    public get verticalAlignment(): VerticalAlignment|undefined {
        return xmlq.getChildAttributeAsString(this.xfNode, 'alignment', 'vertical') as VerticalAlignment|undefined;
    }

    public set verticalAlignment(alignment: VerticalAlignment|undefined) {
        xmlq.setChildAttributes(this.xfNode, 'alignment', { vertical: alignment });
        xmlq.removeChildIfEmpty(this.xfNode, 'alignment');
    }

    /**
     * Gets/sets whether the text in the cell is wrapped.
     */
    public get wrapText(): boolean {
        return xmlq.getChildAttribute(this.xfNode, 'alignment', 'wrapText') === 1;
    }

    public set wrapText(wrapText: boolean) {
        xmlq.setChildAttributes(this.xfNode, 'alignment', { wrapText: wrapText ? 1 : undefined });
        xmlq.removeChildIfEmpty(this.xfNode, 'alignment');
    }

    /**
     * Gets/sets whether the text should be shrunk to fit the cell.
     */
    public get shrinkToFit(): boolean {
        return xmlq.getChildAttribute(this.xfNode, 'alignment', 'shrinkToFit') === 1;
    }

    public set shrinkToFit(shrinkToFit: boolean) {
        xmlq.setChildAttributes(this.xfNode, 'alignment', { shrinkToFit: shrinkToFit ? 1 : undefined });
        xmlq.removeChildIfEmpty(this.xfNode, 'alignment');
    }

    /**
     * Gets/sets the text direction.
     */
    public get textDirection(): TextDirection|undefined {
        const readingOrder = xmlq.getChildAttribute(this.xfNode, 'alignment', 'readingOrder');
        if (readingOrder === 1) return 'left-to-right';
        if (readingOrder === 2) return 'right-to-left';
    }

    public set textDirection(textDirection: TextDirection|undefined) {
        let readingOrder;
        if (textDirection === 'left-to-right') readingOrder = 1;
        else if (textDirection === 'right-to-left') readingOrder = 2;
        xmlq.setChildAttributes(this.xfNode, 'alignment', { readingOrder });
        xmlq.removeChildIfEmpty(this.xfNode, 'alignment');
    }

    /**
     * Rotations in Excel are [-90, 90] but in OOXML are [0,180]
     */
    private get textRotationPositive(): number|undefined {
        return xmlq.getChildAttributeAsNumber(this.xfNode, 'alignment', 'textRotation');
    }

    private set textRotationPositive(textRotation: number|undefined) {
        xmlq.setChildAttributes(this.xfNode, 'alignment', { textRotation });
        xmlq.removeChildIfEmpty(this.xfNode, 'alignment');
    }

    /**
     * Gets/sets the rotation of the text. Counter-clockwise angle of rotation in degrees. Must be [-90, 90] where
     * negative numbers indicate clockwise rotation.
     */
    public get textRotation(): number|undefined {
        let textRotation = this.textRotationPositive;
        if (_.isNil(textRotation)) return;
        if (textRotation > 90) textRotation = 90 - textRotation;
        return textRotation;
    }

    public set textRotation(textRotation: number|undefined) {
        if (_.isNil(textRotation)) {
            this.textRotationPositive = undefined;
        } else {
            if (textRotation < 0) textRotation = 90 - textRotation;
            this.textRotationPositive = textRotation;
        }
    }

    /**
     * Gets/sets whether the text is rotated 45 degrees.
     */
    public get angleTextCounterclockwise(): boolean {
        return this.textRotation === 45;
    }

    public set angleTextCounterclockwise(value: boolean) {
        this.textRotation = value ? 45 : undefined;
    }

    /**
     * Gets/sets whether the text is rotated -45 degrees.
     */
    public get angleTextClockwise(): boolean {
        return this.textRotation === -45;
    }

    public set angleTextClockwise(value: boolean) {
        this.textRotation = value ? -45 : undefined;
    }

    /**
     * Gets/sets whether the text is rotated 90 degrees.
     */
    public get rotateTextUp(): boolean {
        return this.textRotation === 90;
    }

    public set rotateTextUp(value: boolean) {
        this.textRotation = value ? 90 : undefined;
    }

    /**
     * Gets/sets whether the text is rotated -90 degrees.
     */
    public get rotateTextDown(): boolean {
        return this.textRotation === -90;
    }

    public set rotateTextDown(value: boolean) {
        this.textRotation = value ? -90 : undefined;
    }

    /**
     * Gets/sets special rotation that shows text vertical but individual letters are oriented normally.
     */
    public get verticalText(): boolean {
        return this.textRotationPositive === 255;
    }

    public set verticalText(value: boolean) {
        this.textRotationPositive = value ? 255 : undefined;
    }

    public get fill(): Fill|undefined {
        const patternFillNode = xmlq.findChild(this.fillNode, 'patternFill');
        const gradientFillNode = xmlq.findChild(this.fillNode, 'gradientFill');
        
        if (patternFillNode) {
            const patternType = patternFillNode.attributes && patternFillNode.attributes.patternType;

            if (patternType === 'solid') {
                return {
                    type: 'solid',
                    color: getColor(patternFillNode, 'fgColor')!,
                };
            }

            return {
                type: 'pattern',
                pattern: patternType as FillPattern,
                foreground: getColor(patternFillNode, 'fgColor')!,
                background: getColor(patternFillNode, 'bgColor')!,
            };
        }

        if (gradientFillNode) {
            const gradientType = (gradientFillNode.attributes && gradientFillNode.attributes.type) as GradientType || 'linear';
            const fill: GradientFill = {
                type: 'gradient',
                gradientType,
                stops: [],
            };

            _.forEach(gradientFillNode.children, stop => {
                if (typeof stop !== 'string' && typeof stop !== 'number' && stop.attributes) {
                    fill.stops.push({
                        position: Number(stop.attributes.position),
                        color: getColor(stop, 'color')!,
                    });
                }
            });

            if (gradientFillNode.attributes) {
                if (gradientType === 'linear') {
                    fill.angle = Number(gradientFillNode.attributes.degree);
                } else {
                    fill.left = Number(gradientFillNode.attributes.left);
                    fill.right = Number(gradientFillNode.attributes.right);
                    fill.top = Number(gradientFillNode.attributes.top);
                    fill.bottom = Number(gradientFillNode.attributes.bottom);
                }
            }

            return fill;
        }
    }

    public set fill(fill: Fill|undefined) {
        this.fillNode.children = [];

        // No fill
        if (_.isNil(fill)) return;

        // Pattern fill
        if (fill.type === 'pattern') {
            const patternFill = {
                name: 'patternFill',
                attributes: { patternType: fill.pattern },
                children: [],
            };
            this.fillNode.children.push(patternFill);
            setColor(patternFill, 'fgColor', fill.foreground);
            setColor(patternFill, 'bgColor', fill.background);
            return;
        }

        // Gradient fill
        if (fill.type === 'gradient') {
            const gradientFill: INode = { name: 'gradientFill', attributes: {}, children: [] };
            this.fillNode.children.push(gradientFill);
            xmlq.setAttributes(gradientFill, {
                type: fill.gradientType === 'path' ? 'path' : undefined,
                left: fill.left,
                right: fill.right,
                top: fill.top,
                bottom: fill.bottom,
                degree: fill.angle,
            });

            _.forEach(fill.stops, fillStop => {
                const stop = {
                    name: 'stop',
                    attributes: { position: fillStop.position },
                    children: [],
                };

                xmlq.appendChild(gradientFill, stop);
                setColor(stop, 'color', fillStop.color);
            });

            return;
        }

        // Solid fill (really a pattern fill with a solid pattern type).
        if (fill.type === 'solid') {
            const patternFill = {
                name: 'patternFill',
                attributes: { patternType: 'solid' },
            };
            this.fillNode.children.push(patternFill);
            setColor(patternFill, 'fgColor', fill.color);
        }
    }

    public get borders(): Borders {
        return this._borders;
    }

    public get numberFormat(): string {
        const numFmtId = Number(this.xfNode.attributes && this.xfNode.attributes.numFmtId) || 0;
        return this.styleSheet.getNumberFormatCode(numFmtId);
    }

    public set numberFormat(formatCode: string) {
        xmlq.setAttributes(this.xfNode, { numFmtId: this.styleSheet.getNumberFormatId(formatCode) });
    }
}
