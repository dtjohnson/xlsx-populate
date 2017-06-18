"use strict";

const proxyquire = require("proxyquire");
const Promise = require("jszip").external.Promise;

describe("XmlParser", () => {
    let XmlParser, xmlParser, externals;

    beforeEach(() => {
        // proxyquire doesn't like overriding raw objects... a spy obj works.
        externals = jasmine.createSpyObj("externals", ["_"]);
        externals.Promise = Promise;

        XmlParser = proxyquire("../../lib/XmlParser", {
            './externals': externals,
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
    <D xml:space="preserve">    
    </D>
    <E>01</E>
    <F>1</F>
    <G>-1.23</G>
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
                                    { name: 'C', attributes: {}, children: [] },
                                    { name: 'D', attributes: { 'xml:space': "preserve" }, children: ["    \n    "] },
                                    { name: 'E', attributes: {}, children: ["01"] },
                                    { name: 'F', attributes: {}, children: [1] },
                                    { name: 'G', attributes: {}, children: [-1.23] }
                                ]
                            },
                            "bar"
                        ]
                    });
                });
        });
    });
});
