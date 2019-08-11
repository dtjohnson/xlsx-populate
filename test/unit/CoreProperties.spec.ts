import { CoreProperties } from '../../src/CoreProperties';
import { INode } from '../../src/XmlParser';

fdescribe('CoreProperties', () => {
    let coreProperties: CoreProperties, corePropertiesNode: INode;

    beforeEach(() => {
        corePropertiesNode = {
            name: 'Types',
            attributes: {
                xmlns: 'http://schemas.openxmlformats.org/package/2006/content-types'
            },
            children: [],
        };

        coreProperties = new CoreProperties(corePropertiesNode);
    });

    describe('title', () => {
        it('should get/set the title', () => {
            expect(coreProperties.title).toBeUndefined();

            coreProperties.title = 'TITLE';
            expect(coreProperties.title).toBe('TITLE');
            expect(corePropertiesNode.children).toEqual([
                { name: 'dc:title', attributes: {}, children: [ 'TITLE' ] },
            ]);

            coreProperties.title = undefined;
            expect(coreProperties.title).toBeUndefined();
            expect(corePropertiesNode.children).toEqual([]);
        });
    });

    describe('subject', () => {
        it('should get/set the subject', () => {
            expect(coreProperties.subject).toBeUndefined();

            coreProperties.subject = 'SUBJECT';
            expect(coreProperties.subject).toBe('SUBJECT');
            expect(corePropertiesNode.children).toEqual([
                { name: 'dc:subject', attributes: {}, children: [ 'SUBJECT' ] },
            ]);

            coreProperties.subject = undefined;
            expect(coreProperties.subject).toBeUndefined();
            expect(corePropertiesNode.children).toEqual([]);
        });
    });

    describe('author', () => {
        it('should get/set the author', () => {
            expect(coreProperties.author).toBeUndefined();

            coreProperties.author = 'AUTHOR';
            expect(coreProperties.author).toBe('AUTHOR');
            expect(corePropertiesNode.children).toEqual([
                { name: 'dc:creator', attributes: {}, children: [ 'AUTHOR' ] },
            ]);

            coreProperties.author = undefined;
            expect(coreProperties.author).toBeUndefined();
            expect(corePropertiesNode.children).toEqual([]);
        });
    });

    describe('keywords', () => {
        it('should get/set the keywords', () => {
            expect(coreProperties.keywords).toBeUndefined();

            coreProperties.keywords = 'KEYWORDS';
            expect(coreProperties.keywords).toBe('KEYWORDS');
            expect(corePropertiesNode.children).toEqual([
                { name: 'cp:keywords', attributes: {}, children: [ 'KEYWORDS' ] },
            ]);

            coreProperties.keywords = undefined;
            expect(coreProperties.keywords).toBeUndefined();
            expect(corePropertiesNode.children).toEqual([]);
        });
    });

    describe('comments', () => {
        it('should get/set the comments', () => {
            expect(coreProperties.comments).toBeUndefined();

            coreProperties.comments = 'COMMENTS';
            expect(coreProperties.comments).toBe('COMMENTS');
            expect(corePropertiesNode.children).toEqual([
                { name: 'dc:description', attributes: {}, children: [ 'COMMENTS' ] },
            ]);

            coreProperties.comments = undefined;
            expect(coreProperties.comments).toBeUndefined();
            expect(corePropertiesNode.children).toEqual([]);
        });
    });

    describe('lastModifiedBy', () => {
        it('should get/set the lastModifiedBy', () => {
            expect(coreProperties.lastModifiedBy).toBeUndefined();

            coreProperties.lastModifiedBy = 'LAST_MODIFIED_BY';
            expect(coreProperties.lastModifiedBy).toBe('LAST_MODIFIED_BY');
            expect(corePropertiesNode.children).toEqual([
                { name: 'cp:lastModifiedBy', attributes: {}, children: [ 'LAST_MODIFIED_BY' ] },
            ]);

            coreProperties.lastModifiedBy = undefined;
            expect(coreProperties.lastModifiedBy).toBeUndefined();
            expect(corePropertiesNode.children).toEqual([]);
        });
    });

    describe('created', () => {
        it('should get/set the created', () => {
            expect(coreProperties.created).toBeUndefined();

            const date = new Date(2001, 0, 1);
            coreProperties.created = date;
            expect(coreProperties.created).toEqual(date);
            expect(corePropertiesNode.children).toEqual([
                { name: 'dcterms:created', attributes: { 'xsi:type': 'dcterms:W3CDTF' }, children: [ '2001-01-01T05:00:00Z' ] },
            ]);

            coreProperties.created = undefined;
            expect(coreProperties.created).toBeUndefined();
            expect(corePropertiesNode.children).toEqual([]);
        });
    });

    describe('modified', () => {
        it('should get/set the modified', () => {
            expect(coreProperties.modified).toBeUndefined();

            const date = new Date(2001, 0, 1);
            coreProperties.modified = date;
            expect(coreProperties.modified).toEqual(date);
            expect(corePropertiesNode.children).toEqual([
                { name: 'dcterms:modified', attributes: { 'xsi:type': 'dcterms:W3CDTF' }, children: [ '2001-01-01T05:00:00Z' ] },
            ]);

            coreProperties.modified = undefined;
            expect(coreProperties.modified).toBeUndefined();
            expect(corePropertiesNode.children).toEqual([]);
        });
    });

    describe('category', () => {
        it('should get/set the category', () => {
            expect(coreProperties.category).toBeUndefined();

            coreProperties.category = 'CATEGORY';
            expect(coreProperties.category).toBe('CATEGORY');
            expect(corePropertiesNode.children).toEqual([
                { name: 'cp:category', attributes: {}, children: [ 'CATEGORY' ] },
            ]);

            coreProperties.category = undefined;
            expect(coreProperties.category).toBeUndefined();
            expect(corePropertiesNode.children).toEqual([]);
        });
    });
});
