import { Borders } from '../../src/Borders';
import { Border, DiagonalBorder } from '../../src/types';
import { INode } from '../../src/XmlParser';

describe('Borders', () => {
    let borders: Borders, bordersNode: INode;

    beforeEach(() => {
        bordersNode = {
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

        borders = new Borders(bordersNode);
    });

    describe('setBorder', () => {
        it('should set the border on the specified side', () => {
            borders['setBorder']('top', { style: 'thick', color: { rgb: '#ff0000' } });

            expect(bordersNode).toEqual({
                name: 'border',
                attributes: {},
                children: [
                    { name: 'left', attributes: {}, children: [] },
                    { name: 'right', attributes: {}, children: [] },
                    { name: 'top', attributes: { style: 'thick' }, children: [
                        { name: 'color', attributes: { rgb: '#FF0000' }, children: [] },
                    ] },
                    { name: 'bottom', attributes: {}, children: [] },
                    { name: 'diagonal', attributes: {}, children: [] },
                ],
            });

            borders['setBorder']('top', undefined);

            expect(bordersNode).toEqual({
                name: 'border',
                attributes: {},
                children: [
                    { name: 'left', attributes: {}, children: [] },
                    { name: 'right', attributes: {}, children: [] },
                    { name: 'top', attributes: {}, children: [] },
                    { name: 'bottom', attributes: {}, children: [] },
                    { name: 'diagonal', attributes: {}, children: [] },
                ],
            });
        });
    });

    describe('getBorder', () => {
        it('should get the border on the specified side', () => {
            bordersNode.children = [
                { name: 'left', attributes: {}, children: [] },
                { name: 'right', attributes: { style: 'thick' }, children: [
                    { name: 'color', attributes: { rgb: '#FF0000' }, children: [] },
                ] },
                { name: 'top', attributes: {}, children: [] },
                { name: 'bottom', attributes: {}, children: [] },
                { name: 'diagonal', attributes: {}, children: [] },
            ];

            expect(borders['getBorder']('right')).toEqual({
                style: 'thick',
                color: { rgb: '#FF0000' },
            });

            expect(borders['getBorder']('top')).toBeUndefined();
        });
    });

    describe('getter/setters', () => {
        const border: Border = { style: 'thick' };

        beforeEach(() => {
            borders['getBorder'] = jasmine.createSpy('getBorder').and.returnValue(border);
            borders['setBorder'] = jasmine.createSpy('setBorder').and.returnValue(undefined);
        });

        it('should get the top border', () => {
            expect(borders.top).toBe(border);
            expect(borders['getBorder']).toHaveBeenCalledWith('top');
        });

        it('should set the top border', () => {
            borders.top = border;
            expect(borders['setBorder']).toHaveBeenCalledWith('top', border);
        });

        it('should get the bottom border', () => {
            expect(borders.bottom).toBe(border);
            expect(borders['getBorder']).toHaveBeenCalledWith('bottom');
        });

        it('should set the bottom border', () => {
            borders.bottom = border;
            expect(borders['setBorder']).toHaveBeenCalledWith('bottom', border);
        });

        it('should get the left border', () => {
            expect(borders.left).toBe(border);
            expect(borders['getBorder']).toHaveBeenCalledWith('left');
        });

        it('should set the left border', () => {
            borders.left = border;
            expect(borders['setBorder']).toHaveBeenCalledWith('left', border);
        });

        it('should get the right border', () => {
            expect(borders.right).toBe(border);
            expect(borders['getBorder']).toHaveBeenCalledWith('right');
        });

        it('should set the right border', () => {
            borders.right = border;
            expect(borders['setBorder']).toHaveBeenCalledWith('right', border);
        });

        it('should get the diagonal border', () => {
            bordersNode.attributes = {
                diagonalUp: 1,
            };

            expect(borders.diagonal).toEqual({
                style: 'thick',
                direction: 'up',
            });
            expect(borders['getBorder']).toHaveBeenCalledWith('diagonal');

            bordersNode.attributes = {
                diagonalDown: 1,
            };

            expect(borders.diagonal).toEqual({
                style: 'thick',
                direction: 'down',
            });

            bordersNode.attributes = {
                diagonalUp: 1,
                diagonalDown: 1,
            };

            expect(borders.diagonal).toEqual({
                style: 'thick',
                direction: 'both',
            });
        });

        it('should set the diagonal border', () => {
            const diagonal: DiagonalBorder = { style: 'thick', direction: 'both' };
            borders.diagonal = diagonal;
            expect(borders['setBorder']).toHaveBeenCalledWith('diagonal', diagonal);
            expect(bordersNode.attributes).toEqual({
                diagonalUp: 1,
                diagonalDown: 1,
            });
        });
    });
});
