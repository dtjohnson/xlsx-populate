"use strict";

const xmlbuilder = require("xmlbuilder");
const proxyquire = require("proxyquire");

describe("XmlBuilder", () => {
    let XmlBuilder, xmlBuilder;

    beforeEach(() => {
        XmlBuilder = proxyquire("../lib/XmlBuilder", {
            xmlbuilder, // xmlbuilder doesn't play nice with proxyquireify, so include it this way
            '@noCallThru': true
        });
        xmlBuilder = new XmlBuilder();
    });

    describe("build", () => {
        it("should create the XML", () => {
            const node = {
                name: 'root',
                attributes: {
                    foo: 1,
                    bar: "something"
                },
                children: [
                    "foo",
                    {
                        name: 'child',
                        children: [
                            { name: 'A', attributes: {}, children: ["TEXT"] },
                            { name: 'B', attributes: { 'foo:bar': "value" } },
                            { name: 'C' }
                        ]
                    },
                    "bar"
                ]
            };

            expect(xmlBuilder.build(node)).toBe(
`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<root foo="1" bar="something">
  foo
  <child>
    <A>TEXT</A>
    <B foo:bar="value"/>
    <C/>
  </child>
  bar
</root>`);
        });
    });
});
