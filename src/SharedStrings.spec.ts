import { SharedStrings } from './SharedStrings';
import { INode } from './XmlParser';

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
            expect((ss as any)._node).toEqual({
                name: 'sst',
                attributes: {
                    xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
                },
            });
        });

        it('should clear the counts', () => {
            expect((sharedStrings as any)._node.attributes).toEqual({
                xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
            });
        });
    });

    describe('getIndexForString', () => {
        beforeEach(() => {
            (sharedStrings as any)._stringArray = [
                'foo',
                'bar',
                [ { name: 'r' }, { name: 'r' } ],
            ];

            (sharedStrings as any)._indexMap = {
                foo: 0,
                bar: 1,
                '[{"name":"r"},{"name":"r"}]': 2,
            };
        });

        it('should return the index if the string already exists', () => {
            expect(sharedStrings.getIndexForString('foo')).toBe(0);
            expect(sharedStrings.getIndexForString('bar')).toBe(1);
            expect(sharedStrings.getIndexForString([ { name: 'r' }, { name: 'r' } ])).toBe(2);
        });

        it("should create a new entry if the string doesn't exist", () => {
            expect(sharedStrings.getIndexForString('baz')).toBe(3);
            expect((sharedStrings as any)._stringArray).toEqual([
                'foo',
                'bar',
                [ { name: 'r' }, { name: 'r' } ],
                'baz',
            ]);
            expect((sharedStrings as any)._indexMap).toEqual({
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
            expect((sharedStrings as any)._stringArray).toEqual([
                'foo',
                'bar',
                [ { name: 'r' }, { name: 'r' } ],
                [ { name: 'r' } ],
            ]);
            expect((sharedStrings as any)._indexMap).toEqual({
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
            (sharedStrings as any)._stringArray = [ 'foo', 'bar', 'baz' ];
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

    describe('_cacheExistingSharedStrings', () => {
        it('should cache the existing shared strings', () => {
            (sharedStrings as any)._node.children = [
                { name: 'si', children: [ { name: 't', children: [ 'foo' ] } ] },
                { name: 'si', children: [ { name: 't', children: [ 'bar' ] } ] },
                { name: 'si', children: [ { name: 'r', children: [ {} ] }, { name: 'r', children: [ {} ] } ] },
                { name: 'si', children: [ { name: 't', children: [ 'baz' ] } ] },
            ];

            (sharedStrings as any)._stringArray = [];
            (sharedStrings as any)._indexMap = {};
            (sharedStrings as any)._cacheExistingSharedStrings();

            expect((sharedStrings as any)._stringArray).toEqual([
                'foo',
                'bar',
                [ { name: 'r', children: [ {} ] }, { name: 'r', children: [ {} ] } ],
                'baz',
            ]);
            expect((sharedStrings as any)._indexMap).toEqual({
                foo: 0,
                bar: 1,
                '[{"name":"r","children":[{}]},{"name":"r","children":[{}]}]': 2,
                baz: 3,
            });
        });
    });
});
