import { Relationships } from '../../src/Relationships';
import { INode } from '../../src/XmlParser';

describe('Relationships', () => {
    let relationships: Relationships, relationshipsNode: INode;

    beforeEach(() => {
        relationshipsNode = {
            name: 'Relationships',
            attributes: {
                xmlns: 'http://schemas.openxmlformats.org/package/2006/relationships',
            },
            children: [
                {
                    name: 'Relationship',
                    attributes: {
                        Id: 'rId2',
                        Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
                        Target: 'theme/theme1.xml',
                    },
                },
                {
                    name: 'Relationship',
                    attributes: {
                        Id: 'rId1',
                        Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
                        Target: 'worksheets/sheet1.xml',
                    },
                },
            ],
        };

        relationships = new Relationships(relationshipsNode);
    });

    describe('constructor', () => {
        it('should create the node if needed', () => {
            const r = new Relationships();
            expect((r as any).node).toEqual({
                name: 'Relationships',
                attributes: {
                    xmlns: 'http://schemas.openxmlformats.org/package/2006/relationships',
                },
            });
        });

        it('should set the next ID to 1 if no children', () => {
            relationshipsNode.children = [];
            const r = new Relationships(relationshipsNode);
            expect((r as any).nextId).toBe(1);
        });

        it('should set the next ID to last found ID + 1', () => {
            relationshipsNode.children = [
                { name: 'Relationship', attributes: { Id: 'rId2' } },
                { name: 'Relationship', attributes: { Id: 'rId1' } },
                { name: 'Relationship', attributes: { Id: 'rId3' } },
            ];
            const r = new Relationships(relationshipsNode);
            expect((r as any).nextId).toBe(4);
        });
    });

    describe('add', () => {
        it('should add a new relationship', () => {
            relationships.add('TYPE', 'TARGET');
            expect(relationshipsNode.children![2]).toEqual({
                name: 'Relationship',
                attributes: {
                    Id: 'rId3',
                    Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/TYPE',
                    Target: 'TARGET',
                },
            });
        });

        it('should add a new relationship with target mode', () => {
            relationships.add('TYPE', 'TARGET', 'TARGET_MODE');
            expect(relationshipsNode.children![2]).toEqual({
                name: 'Relationship',
                attributes: {
                    Id: 'rId3',
                    Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/TYPE',
                    Target: 'TARGET',
                    TargetMode: 'TARGET_MODE',
                },
            });
        });
    });

    describe('findById', () => {
        it('should return the relationship if matched', () => {
            expect(relationships.findById('rId1')).toBe(relationshipsNode.children![1] as any);
            expect(relationships.findById('rId2')).toBe(relationshipsNode.children![0] as any);
        });

        it('should return undefined if not matched', () => {
            expect(relationships.findById('rId5')).toBeUndefined();
        });
    });

    describe('findByType', () => {
        it('should return the relationship if matched', () => {
            expect(relationships.findByType('worksheet')).toBe(relationshipsNode.children![1] as any);
            expect(relationships.findByType('theme')).toBe(relationshipsNode.children![0] as any);
        });

        it('should return undefined if not matched', () => {
            expect(relationships.findByType('foo')).toBeUndefined();
        });
    });

    describe('toXml', () => {
        it('should return the node as is', () => {
            expect(relationships.toXml()).toBe(relationshipsNode);
        });

        it('should return undefined', () => {
            relationshipsNode.children = [];
            expect(relationships.toXml()).toBeUndefined();
        });
    });
});
