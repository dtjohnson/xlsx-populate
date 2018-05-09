"use strict";

const proxyquire = require("proxyquire");

describe("CoreProperties", () => {
    let CoreProperties, coreProperties, corePropertiesNode;

    beforeEach(() => {
        CoreProperties = proxyquire("../../lib/CoreProperties", {
            '@noCallThru': true
        });

        corePropertiesNode = {
            name: "Types",
            attributes: {
                xmlns: "http://schemas.openxmlformats.org/package/2006/content-types"
            },
            children: []
        };

        coreProperties = new CoreProperties(corePropertiesNode);
    });

    describe("set", () => {
        it("should set a property value", () => {
            coreProperties.set("Title", "A_TITLE");
            expect(coreProperties._properties.title).toBe("A_TITLE");
        });

        it("should throw if not an allowed property name", () => {
            let invalidPropertyName = "invalid-property-name";
            expect(() => {
                coreProperties.set(invalidPropertyName, "SOME_VALUE");
            }).toThrow(new Error(`Unknown property name: "${invalidPropertyName}"`));
        });
    });

    describe("get", () => {
        it("should get a property value", () => {
            coreProperties.set("title", "A_TITLE");
            expect(coreProperties.get("title")).toBe("A_TITLE");
        });

        it("should throw if not an allowed property name", () => {
            let invalidPropertyName = "invalid-property-name";
            expect(() => {
                coreProperties.get(invalidPropertyName);
            }).toThrow(new Error(`Unknown property name: "${invalidPropertyName}"`));
        });
    });

    describe("toXml", () => {
        it("should return the node as is", () => {
            coreProperties.set("Title", "A_TITLE");

            expect(coreProperties.toXml()).toBe(corePropertiesNode);
        });
    });
});
