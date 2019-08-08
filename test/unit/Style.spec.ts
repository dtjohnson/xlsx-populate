import { Style } from '../../src/Style';
import { NumberFormatSource } from '../../src/types';
import { INode } from '../../src/XmlParser';

describe('Style', () => {
    let style: Style, styleSheet: jasmine.SpyObj<NumberFormatSource>, id, xfNode: INode, fontNode: INode, fillNode: INode, borderNode;

    beforeEach(() => {
        styleSheet = jasmine.createSpyObj<NumberFormatSource>('styleSheet', [ 'getNumberFormatCode', 'getNumberFormatId' ]);
        id = 78;
        xfNode = { name: 'xf', attributes: {}, children: [] };
        fontNode = { name: 'font', attributes: {}, children: [] };
        fillNode = { name: 'fill', attributes: {}, children: [] };
        borderNode = {
            name: 'border',
            attributes: {},
            children: [
                { name: 'left', attributes: {}, children: [] },
                { name: 'right', attributes: {}, children: [] },
                { name: 'top', attributes: {}, children: [] },
                { name: 'bottom', attributes: {}, children: [] },
                { name: 'diagonal', attributes: {}, children: [] },
            ],
        };
        
        style = new Style(styleSheet, id, xfNode, fontNode, fillNode, borderNode);
    });

    describe('id', () => {
        it('should return the ID', () => {
            expect(style.id).toBe(78);
        });
    });

    describe('bold', () => {
        it('should get/set bold', () => {
            expect(style.bold).toBe(false);

            style.bold = true;
            expect(style.bold).toBe(true);
            expect(fontNode.children).toEqual([ { name: 'b', attributes: {}, children: [] } ]);

            style.bold = false;
            expect(style.bold).toBe(false);
            expect(fontNode.children).toEqual([]);
        });
    });

    describe('italic', () => {
        it('should get/set italic', () => {
            expect(style.italic).toBe(false);

            style.italic = true;
            expect(style.italic).toBe(true);
            expect(fontNode.children).toEqual([ { name: 'i', attributes: {}, children: [] } ]);

            style.italic = false;
            expect(style.italic).toBe(false);
            expect(fontNode.children).toEqual([]);
        });
    });

    describe('underline', () => {
        it('should get/set underline', () => {
            expect(style.underline).toBe(false);

            style.underline = true;
            expect(style.underline).toBe(true);
            expect(fontNode.children).toEqual([ { name: 'u', attributes: {}, children: [] } ]);

            style.underline = 'double';
            expect(style.underline).toBe('double');
            expect(fontNode.children).toEqual([ { name: 'u', attributes: { val: 'double' }, children: [] } ]);

            style.underline = true;
            expect(style.underline).toBe(true);
            expect(fontNode.children).toEqual([ { name: 'u', attributes: {}, children: [] } ]);

            style.underline = false;
            expect(style.underline).toBe(false);
            expect(fontNode.children).toEqual([]);
        });
    });

    describe('strikethrough', () => {
        it('should get/set strikethrough', () => {
            expect(style.strikethrough).toBe(false);

            style.strikethrough = true;
            expect(style.strikethrough).toBe(true);
            expect(fontNode.children).toEqual([ { name: 'strike', attributes: {}, children: [] } ]);
            
            style.strikethrough = false;
            expect(style.strikethrough).toBe(false);
            expect(fontNode.children).toEqual([]);
        });
    });

    describe('subscript', () => {
        it('should get/set subscript', () => {
            expect(style.subscript).toBe(false);

            style.subscript = true;
            expect(style.subscript).toBe(true);
            expect(fontNode.children).toEqual([ { name: 'vertAlign', attributes: { val: 'subscript' }, children: [] } ]);

            style.subscript = false;
            expect(style.subscript).toBe(false);
            expect(fontNode.children).toEqual([]);
        });
    });

    describe('superscript', () => {
        it('should get/set superscript', () => {
            expect(style.superscript).toBe(false);

            style.superscript = true;
            expect(style.superscript).toBe(true);
            expect(fontNode.children).toEqual([ { name: 'vertAlign', attributes: { val: 'superscript' }, children: [] } ]);

            style.superscript = false;
            expect(style.superscript).toBe(false);
            expect(fontNode.children).toEqual([]);
        });
    });

    describe('fontSize', () => {
        it('should get/set fontSize', () => {
            expect(style.fontSize).toBe(undefined);

            style.fontSize = 17;
            expect(style.fontSize).toBe(17);
            expect(fontNode.children).toEqual([ { name: 'sz', attributes: { val: 17 }, children: [] } ]);

            style.fontSize = undefined;
            expect(style.fontSize).toBe(undefined);
            expect(fontNode.children).toEqual([]);
        });
    });

    describe('fontFamily', () => {
        it('should get/set fontFamily', () => {
            expect(style.fontFamily).toBe(undefined);

            style.fontFamily = 'Comic Sans MS';
            expect(style.fontFamily).toBe('Comic Sans MS');
            expect(fontNode.children).toEqual([ { name: 'name', attributes: { val: 'Comic Sans MS' }, children: [] } ]);

            style.fontFamily = undefined;
            expect(style.fontFamily).toBe(undefined);
            expect(fontNode.children).toEqual([]);
        });
    });

    describe('fontGenericFamily', () => {
        it('should get/set fontGenericFamily', () => {
            expect(style.fontGenericFamily).toBe(undefined);

            style.fontGenericFamily = 'serif';
            expect(style.fontGenericFamily).toBe('serif');
            expect(fontNode.children).toEqual([ { name: 'family', attributes: { val: 1 }, children: [] } ]);

            style.fontGenericFamily = undefined;
            expect(style.fontGenericFamily).toBe(undefined);
            expect(fontNode.children).toEqual([]);
        });
    });

    describe('fontScheme', () => {
        it('should get/set fontScheme', () => {
            expect(style.fontScheme).toBe(undefined);

            style.fontScheme = 'minor';
            expect(style.fontScheme).toBe('minor');
            expect(fontNode.children).toEqual([ { name: 'scheme', attributes: { val: 'minor' }, children: [] } ]);
            
            style.fontScheme = undefined;
            expect(style.fontScheme).toBe(undefined);
            expect(fontNode.children).toEqual([]);
        });
    });

    describe('fontColor', () => {
        it('should get/set fontColor', () => {
            expect(style.fontColor).toBe(undefined);

            style.fontColor = { rgb: 'ff0000' };
            expect(style.fontColor).toEqual({ rgb: 'FF0000' });
            expect(fontNode.children).toEqual([ { name: 'color', attributes: { rgb: 'FF0000' }, children: [] } ]);

            style.fontColor = { theme: 5 };
            expect(style.fontColor).toEqual({ theme: 5 });
            expect(fontNode.children).toEqual([ { name: 'color', attributes: { theme: 5 }, children: [] } ]);

            style.fontColor = { theme: 3, tint: -0.2 };
            expect(style.fontColor).toEqual({ theme: 3, tint: -0.2 });
            expect(fontNode.children).toEqual([ { name: 'color', attributes: { theme: 3, tint: -0.2 }, children: [] } ]);

            style.fontColor = undefined;
            expect(style.fontColor).toBe(undefined);
            expect(fontNode.children).toEqual([]);
        });

        it('should get a indexed fontColor as an rgb', () => {
            fontNode.children = [ { name: 'color', attributes: { indexed: 7 }, children: [] } ];
            expect(style.fontColor).toEqual({ rgb: '00FFFF' });
        });
    });

    describe('horizontalAlignment', () => {
        it('should get/set horizontalAlignment', () => {
            expect(style.horizontalAlignment).toBe(undefined);

            style.horizontalAlignment = 'center';
            expect(style.horizontalAlignment).toBe('center');
            expect(xfNode.children).toEqual([ { name: 'alignment', attributes: { horizontal: 'center' }, children: [] } ]);

            style.horizontalAlignment = undefined;
            expect(style.horizontalAlignment).toBe(undefined);
            expect(xfNode.children).toEqual([]);
        });
    });

    describe('justifyLastLine', () => {
        it('should get/set justifyLastLine', () => {
            expect(style.justifyLastLine).toBe(false);

            style.justifyLastLine = true;
            expect(style.justifyLastLine).toBe(true);
            expect(xfNode.children).toEqual([ { name: 'alignment', attributes: { justifyLastLine: 1 }, children: [] } ]);

            style.justifyLastLine = false;
            expect(style.justifyLastLine).toBe(false);
            expect(xfNode.children).toEqual([]);
        });
    });

    describe('indent', () => {
        it('should get/set indent', () => {
            expect(style.indent).toBe(undefined);

            style.indent = 3;
            expect(style.indent).toBe(3);
            expect(xfNode.children).toEqual([ { name: 'alignment', attributes: { indent: 3 }, children: [] } ]);

            style.indent = undefined;
            expect(style.indent).toBe(undefined);
            expect(xfNode.children).toEqual([]);
        });
    });

    describe('verticalAlignment', () => {
        it('should get/set verticalAlignment', () => {
            expect(style.verticalAlignment).toBe(undefined);

            style.verticalAlignment = 'center';
            expect(style.verticalAlignment).toBe('center');
            expect(xfNode.children).toEqual([ { name: 'alignment', attributes: { vertical: 'center' }, children: [] } ]);

            style.verticalAlignment = undefined;
            expect(style.verticalAlignment).toBe(undefined);
            expect(xfNode.children).toEqual([]);
        });
    });

    describe('wrapText', () => {
        it('should get/set wrapText', () => {
            expect(style.wrapText).toBe(false);

            style.wrapText = true;
            expect(style.wrapText).toBe(true);
            expect(xfNode.children).toEqual([ { name: 'alignment', attributes: { wrapText: 1 }, children: [] } ]);

            style.wrapText = false;
            expect(style.wrapText).toBe(false);
            expect(xfNode.children).toEqual([]);
        });
    });

    describe('shrinkToFit', () => {
        it('should get/set shrinkToFit', () => {
            expect(style.shrinkToFit).toBe(false);

            style.shrinkToFit = true;
            expect(style.shrinkToFit).toBe(true);
            expect(xfNode.children).toEqual([ { name: 'alignment', attributes: { shrinkToFit: 1 }, children: [] } ]);

            style.shrinkToFit = false;
            expect(style.shrinkToFit).toBe(false);
            expect(xfNode.children).toEqual([]);
        });
    });

    describe('textDirection', () => {
        it('should get/set textDirection', () => {
            expect(style.textDirection).toBe(undefined);

            style.textDirection = 'left-to-right';
            expect(style.textDirection).toBe('left-to-right');
            expect(xfNode.children).toEqual([ { name: 'alignment', attributes: { readingOrder: 1 }, children: [] } ]);

            style.textDirection = 'right-to-left';
            expect(style.textDirection).toBe('right-to-left');
            expect(xfNode.children).toEqual([ { name: 'alignment', attributes: { readingOrder: 2 }, children: [] } ]);

            style.textDirection = undefined;
            expect(style.textDirection).toBe(undefined);
            expect(xfNode.children).toEqual([]);
        });
    });

    describe('textRotation', () => {
        it('should get/set indent', () => {
            expect(style.textRotation).toBe(undefined);

            style.textRotation = 15;
            expect(style.textRotation).toBe(15);
            expect(xfNode.children).toEqual([ { name: 'alignment', attributes: { textRotation: 15 }, children: [] } ]);

            style.textRotation = -25;
            expect(style.textRotation).toBe(-25);
            expect(xfNode.children).toEqual([ { name: 'alignment', attributes: { textRotation: 115 }, children: [] } ]);

            style.textRotation = undefined;
            expect(style.textRotation).toBe(undefined);
            expect(xfNode.children).toEqual([]);
        });
    });

    describe('angleTextCounterclockwise', () => {
        it('should get/set angleTextCounterclockwise', () => {
            expect(style.angleTextCounterclockwise).toBe(false);

            style.angleTextCounterclockwise = true;
            expect(style.angleTextCounterclockwise).toBe(true);
            expect(xfNode.children).toEqual([ { name: 'alignment', attributes: { textRotation: 45 }, children: [] } ]);

            style.angleTextCounterclockwise = false;
            expect(style.angleTextCounterclockwise).toBe(false);
            expect(xfNode.children).toEqual([]);
        });
    });

    describe('angleTextClockwise', () => {
        it('should get/set angleTextClockwise', () => {
            expect(style.angleTextClockwise).toBe(false);

            style.angleTextClockwise = true;
            expect(style.angleTextClockwise).toBe(true);
            expect(xfNode.children).toEqual([ { name: 'alignment', attributes: { textRotation: 90 + 45 }, children: [] } ]);

            style.angleTextClockwise = false;
            expect(style.angleTextClockwise).toBe(false);
            expect(xfNode.children).toEqual([]);
        });
    });

    describe('rotateTextUp', () => {
        it('should get/set rotateTextUp', () => {
            expect(style.rotateTextUp).toBe(false);

            style.rotateTextUp = true;
            expect(style.rotateTextUp).toBe(true);
            expect(xfNode.children).toEqual([ { name: 'alignment', attributes: { textRotation: 90 }, children: [] } ]);

            style.rotateTextUp = false;
            expect(style.rotateTextUp).toBe(false);
            expect(xfNode.children).toEqual([]);
        });
    });

    describe('rotateTextDown', () => {
        it('should get/set rotateTextDown', () => {
            expect(style.rotateTextDown).toBe(false);

            style.rotateTextDown = true;
            expect(style.rotateTextDown).toBe(true);
            expect(xfNode.children).toEqual([ { name: 'alignment', attributes: { textRotation: 90 + 90 }, children: [] } ]);

            style.rotateTextDown = false;
            expect(style.rotateTextDown).toBe(false);
            expect(xfNode.children).toEqual([]);
        });
    });

    describe('verticalText', () => {
        it('should get/set verticalText', () => {
            expect(style.verticalText).toBe(false);

            style.verticalText = true;
            expect(style.verticalText).toBe(true);
            expect(xfNode.children).toEqual([ { name: 'alignment', attributes: { textRotation: 255 }, children: [] } ]);

            style.verticalText = false;
            expect(style.verticalText).toBe(false);
            expect(xfNode.children).toEqual([]);
        });
    });

    describe('borders', () => {
        it('should get the borders object', () => {
            expect(style.borders).toBe(style['_borders']);
        });
    });

    describe('fill', () => {
        it('should get/set solid fill', () => {
            expect(style.fill).toBe(undefined);

            style.fill = { type: 'solid', color: { rgb: 'ff0000' } };
            expect(style.fill).toEqual({
                type: 'solid',
                color: { rgb: 'FF0000' },
            });
            expect(fillNode.children).toEqual([ {
                name: 'patternFill',
                attributes: { patternType: 'solid' },
                children: [ {
                    name: 'fgColor',
                    attributes: { rgb: 'FF0000' },
                    children: [],
                } ],
            } ]);

            style.fill = { type: 'solid', color: { theme: 6, tint: -0.25 } };
            expect(style.fill).toEqual({
                type: 'solid',
                color: { theme: 6, tint: -0.25 },
            });
            expect(fillNode.children).toEqual([ {
                name: 'patternFill',
                attributes: { patternType: 'solid' },
                children: [ {
                    name: 'fgColor',
                    attributes: { theme: 6, tint: -0.25 },
                    children: [],
                } ],
            } ]);

            style.fill = undefined;
            expect(style.fill).toBe(undefined);
            expect(fillNode.children).toEqual([]);
        });

        it('should get/set pattern fill', () => {
            expect(style.fill).toBe(undefined);

            style.fill = {
                type: 'pattern',
                pattern: 'gray0625',
                foreground: { rgb: 'aa0000', tint: -1 },
                background: { theme: 3, tint: 1 },
            };
            expect(style.fill).toEqual({
                type: 'pattern',
                pattern: 'gray0625',
                foreground: {
                    rgb: 'AA0000',
                    tint: -1,
                },
                background: {
                    theme: 3,
                    tint: 1,
                },
            });
            expect(fillNode.children).toEqual([ {
                name: 'patternFill',
                attributes: { patternType: 'gray0625' },
                children: [ {
                    name: 'fgColor',
                    attributes: { rgb: 'AA0000', tint: -1 },
                    children: [],
                }, {
                    name: 'bgColor',
                    attributes: { theme: 3, tint: 1 },
                    children: [],
                } ],
            } ]);

            style.fill = undefined;
            expect(style.fill).toBe(undefined);
            expect(fillNode.children).toEqual([]);
        });

        it('should get/set gradient fill', () => {
            expect(style.fill).toBe(undefined);

            style.fill = {
                type: 'gradient',
                gradientType: 'path',
                top: 0.1,
                bottom: 0.2,
                left: 0.3,
                right: 0.4,
                stops: [
                    { position: 0, color: { theme: 0, tint: -0.3 } },
                    { position: 1, color: { rgb: 'acacac' } },
                ],
            };
            expect(style.fill).toEqual({
                type: 'gradient',
                gradientType: 'path',
                top: 0.1,
                bottom: 0.2,
                left: 0.3,
                right: 0.4,
                stops: [
                    { position: 0, color: { theme: 0, tint: -0.3 } },
                    { position: 1, color: { rgb: 'ACACAC' } },
                ],
            });
            expect(fillNode.children).toEqual([ {
                name: 'gradientFill',
                attributes: {
                    type: 'path',
                    top: 0.1,
                    bottom: 0.2,
                    left: 0.3,
                    right: 0.4,
                },
                children: [
                    {
                        name: 'stop',
                        attributes: { position: 0 },
                        children: [ { name: 'color', attributes: { theme: 0, tint: -0.3 }, children: [] } ],
                    },
                    {
                        name: 'stop',
                        attributes: { position: 1 },
                        children: [ { name: 'color', attributes: { rgb: 'ACACAC' }, children: [] } ],
                    },
                ],
            } ]);

            style.fill = undefined;
            expect(style.fill).toBe(undefined);
            expect(fillNode.children).toEqual([]);
        });
    });

    describe('numberFormat', () => {
        it('should get/set numberFormat', () => {
            styleSheet.getNumberFormatCode.and.returnValue('foo');
            styleSheet.getNumberFormatId.and.returnValue(7);

            expect(style.numberFormat).toBe('foo');
            expect(styleSheet.getNumberFormatCode).toHaveBeenCalledWith(0);

            style.numberFormat = 'bar';
            expect(styleSheet.getNumberFormatId).toHaveBeenCalledWith('bar');
            expect(xfNode).toEqual({ name: 'xf', attributes: { numFmtId: 7 }, children: [] });
            expect(style.numberFormat).toBe('foo');
            expect(styleSheet.getNumberFormatCode).toHaveBeenCalledWith(7);
        });
    });
});
