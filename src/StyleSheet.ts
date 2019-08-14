
import _ from 'lodash';
import { Style } from './Style';
import { INode } from './XmlParser';
import * as xmlq from './xmlq';

/**
 * Standard number format codes
 * Taken from http://polymathprogrammer.com/2011/02/15/built-in-styles-for-excel-open-xml/
 */
const STANDARD_CODES: { [id: number]: string } = {
    0: 'General',
    1: '0',
    2: '0.00',
    3: '#,##0',
    4: '#,##0.00',
    9: '0%',
    10: '0.00%',
    11: '0.00E+00',
    12: '# ?/?',
    13: '# ??/??',
    14: 'mm-dd-yy',
    15: 'd-mmm-yy',
    16: 'd-mmm',
    17: 'mmm-yy',
    18: 'h:mm AM/PM',
    19: 'h:mm:ss AM/PM',
    20: 'h:mm',
    21: 'h:mm:ss',
    22: 'm/d/yy h:mm',
    37: '#,##0 ;(#,##0)',
    38: '#,##0 ;[Red](#,##0)',
    39: '#,##0.00;(#,##0.00)',
    40: '#,##0.00;[Red](#,##0.00)',
    45: 'mm:ss',
    46: '[h]:mm:ss',
    47: 'mmss.0',
    48: '##0.0E+0',
    49: '@',
};

/**
 * The starting ID for custom number formats. The first 163 indexes are reserved.
 */
const STARTING_CUSTOM_NUMBER_FORMAT_ID = 164;

/**
 * A style sheet.
 */
export class StyleSheet {
    private static readonly Style = Style;

    private readonly numFmtsNode: INode;
    private readonly fontsNode: INode;
    private readonly fillsNode: INode;
    private readonly bordersNode: INode;
    private readonly cellXfsNode: INode;

    private readonly numberFormatCodesById: { [id: number]: string } = {};
    private readonly numberFormatIdsByCode: { [code: string]: number|undefined } = {};

    private nextNumFormatId = STARTING_CUSTOM_NUMBER_FORMAT_ID;

    /**
     * Creates an instance of _StyleSheet.
     * @param node - The style sheet node
     */
    public constructor(private readonly node: INode) {
        // Cache the refs to the collections. The number formats node might not exist.
        const numFmtsNode = xmlq.findChild(this.node, 'numFmts');

        // These should all be present.
        this.fontsNode = xmlq.findChild(this.node, 'fonts')!;
        this.fillsNode = xmlq.findChild(this.node, 'fills')!;
        this.bordersNode = xmlq.findChild(this.node, 'borders')!;
        this.cellXfsNode = xmlq.findChild(this.node, 'cellXfs')!;

        if (numFmtsNode) {
            this.numFmtsNode = numFmtsNode;
        } else {
            this.numFmtsNode = {
                name: 'numFmts',
                attributes: {},
                children: [],
            };

            // Number formats need to be before the others.
            xmlq.insertBefore(this.node, this.numFmtsNode, this.fontsNode);
        }

        // Remove the optional counts so we don't have to keep them up to date.
        xmlq.setAttributes(this.numFmtsNode, { count: undefined });
        xmlq.setAttributes(this.fontsNode, { count: undefined });
        xmlq.setAttributes(this.fillsNode, { count: undefined });
        xmlq.setAttributes(this.bordersNode, { count: undefined });
        xmlq.setAttributes(this.cellXfsNode, { count: undefined });

        this.cacheNumberFormats();
    }

    /**
     * Create a style.
     * @param sourceId - The source style ID to copy, if provided.
     * @returns The style.
     */
    public createStyle(sourceId?: number): Style {
        let fontNode: INode|undefined, fillNode: INode|undefined, borderNode: INode|undefined, xfNode: INode|undefined;
        if (!_.isNil(sourceId)) {
            const sourceXfNode = this.cellXfsNode.children![sourceId] as INode;
            xfNode = _.cloneDeep(sourceXfNode);

            if (sourceXfNode.attributes && sourceXfNode.attributes.applyFont) {
                const fontId = Number(sourceXfNode.attributes.fontId);
                fontNode = _.cloneDeep(this.fontsNode.children![fontId] as INode);
            }

            if (sourceXfNode.attributes && sourceXfNode.attributes.applyFill) {
                const fillId = Number(sourceXfNode.attributes.fillId);
                fillNode = _.cloneDeep(this.fillsNode.children![fillId] as INode);
            }

            if (sourceXfNode.attributes && sourceXfNode.attributes.applyBorder) {
                const borderId = Number(sourceXfNode.attributes.borderId);
                borderNode = _.cloneDeep(this.bordersNode.children![borderId] as INode);
            }
        }

        if (!fontNode) fontNode = { name: 'font', attributes: {}, children: [] };
        xmlq.appendChild(this.fontsNode, fontNode);

        if (!fillNode) fillNode = { name: 'fill', attributes: {}, children: [] };
        xmlq.appendChild(this.fillsNode, fillNode);

        // The border sides must be in order
        if (!borderNode) borderNode = { name: 'border', attributes: {}, children: [] };
        borderNode.children = [
            xmlq.findChild(borderNode, 'left') || { name: 'left', attributes: {}, children: [] },
            xmlq.findChild(borderNode, 'right') || { name: 'right', attributes: {}, children: [] },
            xmlq.findChild(borderNode, 'top') || { name: 'top', attributes: {}, children: [] },
            xmlq.findChild(borderNode, 'bottom') || { name: 'bottom', attributes: {}, children: [] },
            xmlq.findChild(borderNode, 'diagonal') || { name: 'diagonal', attributes: {}, children: [] },
        ];
        xmlq.appendChild(this.bordersNode, borderNode);

        if (!xfNode) xfNode = { name: 'xf', attributes: {}, children: [] };
        _.assign(xfNode.attributes, {
            fontId: this.fontsNode.children!.length - 1,
            fillId: this.fillsNode.children!.length - 1,
            borderId: this.bordersNode.children!.length - 1,
            applyFont: 1,
            applyFill: 1,
            applyBorder: 1,
        });
        xmlq.appendChild(this.cellXfsNode, xfNode);

        const styleId = this.cellXfsNode.children!.length - 1;
        return new StyleSheet.Style(this, styleId, xfNode, fontNode, fillNode, borderNode);
    }

    /**
     * Get the number format code for a given ID.
     * @param id - The number format ID.
     * @returns The format code.
     */
    public getNumberFormatCode(id: number): string {
        return this.numberFormatCodesById[id];
    }

    /**
     * Get the nuumber format ID for a given code.
     * @param code - The format code.
     * @returns The number format ID.
     */
    public getNumberFormatId(code: string): number {
        let id = this.numberFormatIdsByCode[code];
        if (id === undefined) {
            id = this.nextNumFormatId++;
            this.numberFormatCodesById[id] = code;
            this.numberFormatIdsByCode[code] = id;

            xmlq.appendChild(this.numFmtsNode, {
                name: 'numFmt',
                attributes: {
                    numFmtId: id,
                    formatCode: code,
                },
            });
        }

        return id;
    }

    /**
     * Convert the style sheet to an XML object.
     * @returns The XML form.
     */
    public toXml(): INode {
        return this.node;
    }

    /**
     * Cache the number format codes
     */
    private cacheNumberFormats(): void {
        // Load the standard number format codes into the caches.
        for (const id in STANDARD_CODES) {
            if (!(id in STANDARD_CODES)) continue;
            const code = STANDARD_CODES[id];
            this.numberFormatCodesById[id] = code;
            this.numberFormatIdsByCode[code] = parseInt(id, 10);
        }

        // If there are custom number formats, cache them all and update the next num as needed.
        (this.numFmtsNode.children || []).forEach(node => {
            if (typeof node !== 'string' && typeof node !== 'number' && node.attributes) {
                const id = Number(node.attributes.numFmtId);
                const code = String(node.attributes.formatCode);
                this.numberFormatCodesById[id] = code;
                this.numberFormatIdsByCode[code] = id;
                if (id >= this.nextNumFormatId) this.nextNumFormatId = id + 1;
            }
        });
    }
}

// tslint:disable
/*
xl/styles.xml

<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main">
    <numFmts count="1">
        <numFmt numFmtId="164" formatCode="#,##0_);[Red]\(#,##0\)\)"/>
    </numFmts>
    <fonts count="1" x14ac:knownFonts="1">
        <font>
            <sz val="11"/>
            <color theme="1"/>
            <name val="Calibri"/>
            <family val="2"/>
            <scheme val="minor"/>
        </font>
    </fonts>
    <fills count="11">
        <fill>
            <patternFill patternType="none"/>
        </fill>
        <fill>
            <patternFill patternType="gray125"/>
        </fill>
        <fill>
            <patternFill patternType="solid">
                <fgColor rgb="FFC00000"/>
                <bgColor indexed="64"/>
            </patternFill>
        </fill>
        <fill>
            <patternFill patternType="lightDown">
                <fgColor theme="4"/>
                <bgColor rgb="FFC00000"/>
            </patternFill>
        </fill>
        <fill>
            <gradientFill degree="90">
                <stop position="0">
                    <color theme="0"/>
                </stop>
                <stop position="1">
                    <color theme="4"/>
                </stop>
            </gradientFill>
        </fill>
        <fill>
            <gradientFill>
                <stop position="0">
                    <color theme="0"/>
                </stop>
                <stop position="1">
                    <color theme="4"/>
                </stop>
            </gradientFill>
        </fill>
        <fill>
            <gradientFill degree="45">
                <stop position="0">
                    <color theme="0"/>
                </stop>
                <stop position="1">
                    <color theme="4"/>
                </stop>
            </gradientFill>
        </fill>
        <fill>
            <gradientFill degree="135">
                <stop position="0">
                    <color theme="0"/>
                </stop>
                <stop position="1">
                    <color theme="4"/>
                </stop>
            </gradientFill>
        </fill>
        <fill>
            <gradientFill type="path">
                <stop position="0">
                    <color theme="0"/>
                </stop>
                <stop position="1">
                    <color theme="4"/>
                </stop>
            </gradientFill>
        </fill>
        <fill>
            <gradientFill type="path" left="0.5" right="0.5" top="0.5" bottom="0.5">
                <stop position="0">
                    <color theme="0"/>
                </stop>
                <stop position="1">
                    <color theme="4"/>
                </stop>
            </gradientFill>
        </fill>
        <fill>
            <gradientFill degree="270">
                <stop position="0">
                    <color theme="0"/>
                </stop>
                <stop position="1">
                    <color theme="4"/>
                </stop>
            </gradientFill>
        </fill>
    </fills>
    <borders count="10">
        <border>
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal/>
        </border>
        <border diagonalDown="1">
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal style="hair">
                <color auto="1"/>
            </diagonal>
        </border>
        <border diagonalDown="1">
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal style="dotted">
                <color auto="1"/>
            </diagonal>
        </border>
        <border diagonalDown="1">
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal style="dashDotDot">
                <color auto="1"/>
            </diagonal>
        </border>
        <border diagonalDown="1">
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal style="dashDot">
                <color auto="1"/>
            </diagonal>
        </border>
        <border diagonalDown="1">
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal style="dashed">
                <color auto="1"/>
            </diagonal>
        </border>
        <border diagonalUp="1">
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal style="mediumDashDotDot">
                <color auto="1"/>
            </diagonal>
        </border>
        <border diagonalUp="1">
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal style="slantDashDot">
                <color auto="1"/>
            </diagonal>
        </border>
        <border diagonalUp="1">
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal style="mediumDashDot">
                <color auto="1"/>
            </diagonal>
        </border>
        <border diagonalUp="1">
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal style="mediumDashed">
                <color auto="1"/>
            </diagonal>
        </border>
    </borders>
    <cellStyleXfs count="1">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    </cellStyleXfs>
    <cellXfs count="19">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="2" xfId="0" applyBorder="1"/>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="3" xfId="0" applyBorder="1"/>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="4" xfId="0" applyBorder="1"/>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="5" xfId="0" applyBorder="1"/>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="6" xfId="0" applyBorder="1"/>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="7" xfId="0" applyBorder="1"/>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="8" xfId="0" applyBorder="1"/>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="9" xfId="0" applyBorder="1"/>
        <xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1" applyBorder="1"/>
        <xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/>
        <xf numFmtId="0" fontId="0" fillId="4" borderId="0" xfId="0" applyFill="1"/>
        <xf numFmtId="0" fontId="0" fillId="5" borderId="0" xfId="0" applyFill="1"/>
        <xf numFmtId="0" fontId="0" fillId="6" borderId="0" xfId="0" applyFill="1"/>
        <xf numFmtId="0" fontId="0" fillId="7" borderId="0" xfId="0" applyFill="1"/>
        <xf numFmtId="0" fontId="0" fillId="8" borderId="0" xfId="0" applyFill="1"/>
        <xf numFmtId="0" fontId="0" fillId="9" borderId="0" xfId="0" applyFill="1"/>
        <xf numFmtId="0" fontId="0" fillId="10" borderId="0" xfId="0" applyFill="1"/>
    </cellXfs>
    <cellStyles count="1">
        <cellStyle name="Normal" xfId="0" builtinId="0"/>
    </cellStyles>
    <dxfs count="0"/>
    <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
    <extLst>
        <ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
            <x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/>
        </ext>
        <ext uri="{9260A510-F301-46a8-8635-F512D64BE5F5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
            <x15:timelineStyles defaultTimelineStyle="TimeSlicerStyleLight1"/>
        </ext>
    </extLst>
</styleSheet>
*/
