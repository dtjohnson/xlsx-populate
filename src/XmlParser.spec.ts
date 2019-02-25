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
                        attributes: {},
                        children: [
                            { name: 'A', attributes: {}, children: [ 'TEXT' ] },
                            { name: 'B', attributes: { 'foo:bar': 'value' }, children: [] },
                            { name: 'C', attributes: {}, children: [] },
                            { name: 'D', attributes: { 'xml:space': 'preserve' }, children: [ '    \n    ' ] },
                            { name: 'E', attributes: {}, children: [ '01' ] },
                            { name: 'F', attributes: {}, children: [ 1 ] },
                            { name: 'G', attributes: {}, children: [ -1.23 ] },
                        ],
                    },
                    'bar',
                ],
            });
        });
    });
});
