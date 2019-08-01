import { PageBreaks } from '../../src/PageBreaks';
import { INode } from '../../src/XmlParser';

describe('PageBreaks', () => {
    let pageBreaks: PageBreaks, pageBreaksNode: INode;

    beforeEach(() => {
        pageBreaksNode = {
            name: 'rowBreaks',
            attributes: {
                count: 1,
                manualBreakCount: 1,
            },
            children: [
                {
                    name: 'brk',
                    attributes: {
                        id: 8,
                        max: 16383,
                        man: 1,
                    },
                },
            ],
        };

        pageBreaks = new PageBreaks(true, pageBreaksNode);
    });

    describe('constructor', () => {
        it('should create the row breaks node if needed', () => {
            const pb = new PageBreaks(true);
            expect(pb['node']).toEqual({
                name: 'rowBreaks',
            });
        });

        it('should create the col breaks node if needed', () => {
            const pb = new PageBreaks(false);
            expect(pb['node']).toEqual({
                name: 'colBreaks',
            });
        });
    });

    describe('add', () => {
        it('should add a new break', () => {
            pageBreaks.add(7);
            expect(pageBreaksNode.children![1]).toEqual({
                name: 'brk',
                attributes: {
                    id: 7,
                    max: 16383,
                    man: 1,
                },
            });
            expect(pageBreaksNode.attributes).toEqual({
                count: 2,
                manualBreakCount: 2,
            });
        });
    });

    describe('remove', () => {
        it('should remove a break', () => {
            pageBreaks.remove(8);
            expect(pageBreaksNode).toEqual({
                name: 'rowBreaks',
                attributes: {
                    count: 0,
                    manualBreakCount: 0,
                },
                children: [],
            });
        });
    });

    describe('list', () => {
        it('should return the number of breaks', () => {
            expect(pageBreaks.list()).toEqual([ 8 ]);
            pageBreaks.add(9);
            expect(pageBreaks.list()).toEqual([ 8, 9 ]);
        });
    });

    describe('toXml', () => {
        it('should return the node as is', () => {
            expect(pageBreaks.toXml()).toBe(pageBreaksNode);
        });

        it('should return undefined', () => {
            pageBreaksNode.children = [];
            expect(pageBreaks.toXml()).toBeUndefined();
        });
    });
});
