import { XmlParser } from './XmlParser';

describe('XmlParser', () => {
    let xmlParser: XmlParser;

    beforeEach(() => {
        xmlParser = new XmlParser();
    });

    describe('build', () => {
        it('should create the XML', async () => {
            const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<root foo="1" bar="something">foo<child>
    <A>TEXT</A>
    <B foo:bar="value"/>
    <C/>
    <D xml:space="preserve">    
    </D>
    <E>01</E>
    <F>1</F>
    <G>-1.23</G>
  </child>bar</root>`;

            const node = await xmlParser.parseAsync(xml);
            expect(node).toEqual({
                name: 'root',
                attributes: {
                    foo: 1,
                    bar: 'something',
                },
                children: [
                    'foo',
                    {
                        name: 'child',
                        children: [
                            { name: 'A', children: [ 'TEXT' ] },
                            { name: 'B', attributes: { 'foo:bar': 'value' } },
                            { name: 'C' },
                            { name: 'D', attributes: { 'xml:space': 'preserve' }, children: [ '    \n    ' ] },
                            { name: 'E', children: [ '01' ] },
                            { name: 'F', children: [ 1 ] },
                            { name: 'G', children: [ -1.23 ] },
                        ],
                    },
                    'bar',
                ],
            });
        });
    });
});
