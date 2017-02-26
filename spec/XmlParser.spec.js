"use strict";

const proxyquire = require("proxyquire");

describe("XmlParser", () => {
    let XmlParser, xmlParser;

    beforeEach(() => {
        XmlParser = proxyquire("../lib/XmlParser", {
            '@noCallThru': true
        });
        xmlParser = new XmlParser();
    });

    describe("build", () => {
        itAsync("should create the XML", () => {
            const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<root foo="1" bar="something">foo<child>
    <A>TEXT</A>
    <B foo:bar="value"/>
    <C/>
  </child>bar</root>`;

            return xmlParser.parseAsync(xml)
                .then(node => {
                    expect(node).toEqualJson({
                        name: 'root',
                        attributes: {
                            foo: 1,
                            bar: "something"
                        },
                        children: [
                            "foo",
                            {
                                name: 'child',
                                attributes: {},
                                children: [
                                    { name: 'A', attributes: {}, children: ["TEXT"] },
                                    { name: 'B', attributes: { 'foo:bar': "value" }, children: [] },
                                    { name: 'C', attributes: {}, children: [] }
                                ]
                            },
                            "bar"
                        ]
                    });
                });
        });
    });
});
