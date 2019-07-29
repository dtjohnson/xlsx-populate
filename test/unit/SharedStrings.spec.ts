import { SharedStrings } from '../../src/SharedStrings';
import { INode } from '../../src/XmlParser';

describe('SharedStrings', () => {
    let sharedStrings: SharedStrings, sharedStringsNode: INode;

    beforeEach(() => {
        sharedStringsNode = {
            name: 'sst',
            attributes: {
                xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
                count: 3,
                uniqueCount: 7,
            },
            children: [
                {
                    name: 'si',
                    children: [
                        {
                            name: 't',
                            children: [ 'foo' ],
                        },
                    ],
                },
            ],
        };

        sharedStrings = new SharedStrings(sharedStringsNode);
    });

    describe('constructor', () => {
        it('should create the node if needed', () => {
            const ss = new SharedStrings();
            expect((ss as any).node).toEqual({
                name: 'sst',
                attributes: {
                    xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
                },
            });
        });

        it('should clear the counts', () => {
            expect(sharedStrings['node'].attributes).toEqual({
                xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
            });
        });
    });

    describe('getIndexForString', () => {
        beforeEach(() => {
            sharedStrings['stringArray'].length = 0;
            sharedStrings['stringArray'].push(
                'foo',
                'bar',
                [ { name: 'r' }, { name: 'r' } ],
            );

            sharedStrings['indexMap'].foo = 0;
            sharedStrings['indexMap'].bar = 1;
            sharedStrings['indexMap']['[{"name":"r"},{"name":"r"}]'] = 2;
        });

        it('should return the index if the string already exists', () => {
            expect(sharedStrings.getIndexForString('foo')).toBe(0);
            expect(sharedStrings.getIndexForString('bar')).toBe(1);
            expect(sharedStrings.getIndexForString([ { name: 'r' }, { name: 'r' } ])).toBe(2);
        });

        it("should create a new entry if the string doesn't exist", () => {
            expect(sharedStrings.getIndexForString('baz')).toBe(3);
            expect(sharedStrings['stringArray']).toEqual([
                'foo',
                'bar',
                [ { name: 'r' }, { name: 'r' } ],
                'baz',
            ]);
            expect(sharedStrings['indexMap']).toEqual({
                foo: 0,
                bar: 1,
                '[{"name":"r"},{"name":"r"}]': 2,
                baz: 3,
            });
            expect(sharedStringsNode.children![sharedStringsNode.children!.length - 1]).toEqual({
                name: 'si',
                children: [
                    {
                        name: 't',
                        attributes: { 'xml:space': 'preserve' },
                        children: [ 'baz' ],
                    },
                ],
            });
        });

        it("should create a new array entry if the array doesn't exist", () => {
            expect(sharedStrings.getIndexForString([ { name: 'r' } ])).toBe(3);
            expect(sharedStrings['stringArray']).toEqual([
                'foo',
                'bar',
                [ { name: 'r' }, { name: 'r' } ],
                [ { name: 'r' } ],
            ]);
            expect(sharedStrings['indexMap']).toEqual({
                foo: 0,
                bar: 1,
                '[{"name":"r"},{"name":"r"}]': 2,
                '[{"name":"r"}]': 3,
            });
            expect(sharedStringsNode.children![sharedStringsNode.children!.length - 1]).toEqual({
                name: 'si',
                children: [ { name: 'r' } ],
            });
        });
    });

    describe('getStringByIndex', () => {
        it('should return the string at a given index', () => {
            sharedStrings['stringArray'].length = 0;
            sharedStrings['stringArray'].push('foo', 'bar', 'baz');
            expect(sharedStrings.getStringByIndex(0)).toBe('foo');
            expect(sharedStrings.getStringByIndex(1)).toBe('bar');
            expect(sharedStrings.getStringByIndex(2)).toBe('baz');
            expect(sharedStrings.getStringByIndex(3)).toBeUndefined();
        });
    });

    describe('toXml', () => {
        it('should return the node as is', () => {
            expect(sharedStrings.toXml()).toBe(sharedStringsNode);
        });
    });

    describe('cacheExistingSharedStrings', () => {
        it('should cache the existing shared strings', () => {
            sharedStrings['node'].children = [
                { name: 'si', children: [ { name: 't', children: [ 'foo' ] } ] },
                { name: 'si', children: [ { name: 't', children: [ 'bar' ] } ] },
                { name: 'si', children: [ { name: 'r', children: [ { name: 'x' } ] }, { name: 'r', children: [ { name: 'y' } ] } ] },
                { name: 'si', children: [ { name: 't', children: [ 'baz' ] } ] },
            ];

            sharedStrings['stringArray'].length = 0;
            sharedStrings['cacheExistingSharedStrings']();

            expect(sharedStrings['stringArray']).toEqual([
                'foo',
                'bar',
                [ { name: 'r', children: [ { name: 'x' } ] }, { name: 'r', children: [ { name: 'y' } ] } ],
                'baz',
            ]);
            expect(sharedStrings['indexMap']).toEqual({
                foo: 0,
                bar: 1,
                '[{"name":"r","children":[{"name":"x"}]},{"name":"r","children":[{"name":"y"}]}]': 2,
                baz: 3,
            });
        });
    });
});
