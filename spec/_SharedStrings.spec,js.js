"use strict";

const proxyquire = require("proxyquire").noCallThru();
const DOMParser = require('xmldom').DOMParser;
const parser = new DOMParser();

describe("_SharedStrings", () => {
    let _SharedStrings, sharedStrings, sharedStringsText;

    beforeEach(() => {
        _SharedStrings = proxyquire("../lib/_SharedStrings", {});
        sharedStrings = new _SharedStrings();
    });

    describe("constructor", () => {
        it("should create an XML doc if no text passed in", () => {
            expect(sharedStrings._xml.toString()).toBe(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>`);
            expect(sharedStrings._stringArray).toEqual([]);
            expect(sharedStrings._indexMap).toEqual({});
        });

        it("should remove the counts and cache the values", () => {
            sharedStrings = new _SharedStrings(`
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="13" uniqueCount="4">
	<si>
		<t>Foo</t>
	</si>
	<si>
		<r>
			<t>s</t>
		</r>
	</si>
	<si>
		<t>Bar</t>
	</si>
</sst>
`);

            expect(sharedStrings._xml.documentElement.hasAttribute("count")).toBe(false);
            expect(sharedStrings._xml.documentElement.hasAttribute("uniqueCount")).toBe(false);
            expect(sharedStrings._stringArray).toEqual(["Foo", null, "Bar"]);
            expect(sharedStrings._indexMap).toEqual({ Foo: 0, Bar: 2 });
        });
    });

    describe("getIndexForString", () => {
        beforeEach(() => {
            sharedStrings._stringArray = ["foo", "bar"];
            sharedStrings._indexMap = { foo: 0, bar: 1 };
        });

        it("should return the index if the string already exists", () => {
            expect(sharedStrings.getIndexForString("foo")).toBe(0);
            expect(sharedStrings.getIndexForString("bar")).toBe(1);
        });

        it("should create a new entry if the string doesn't exist", () => {
            expect(sharedStrings.getIndexForString("baz")).toBe(2);
            expect(sharedStrings._stringArray).toEqual(["foo", "bar", "baz"]);
            expect(sharedStrings._indexMap).toEqual({ foo: 0, bar: 1, baz: 2 });
            expect(sharedStrings._xml.toString()).toBe(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>baz</t></si></sst>`);
        });
    });

    describe("getStringByIndex", () => {
        it("should return the string at a given index", () => {
            sharedStrings._stringArray = ["foo", "bar", "baz"];
            expect(sharedStrings.getStringByIndex(0)).toBe("foo");
            expect(sharedStrings.getStringByIndex(1)).toBe("bar");
            expect(sharedStrings.getStringByIndex(2)).toBe("baz");
            expect(sharedStrings.getStringByIndex(3)).toBeUndefined();
        });
    });

    describe("toString", () => {
        it("should export to the XML string", () => {
            expect(sharedStrings.toString().trim()).toBe(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>`);
        });
    });
});
