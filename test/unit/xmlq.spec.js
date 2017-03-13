"use strict";

const proxyquire = require("proxyquire");

describe("xmlq", () => {
    let xmlq;

    beforeEach(() => {
        xmlq = proxyquire("../../lib/xmlq", {
            '@noCallThru': true
        });
    });

    describe("appendChild", () => {
        it("should append the child", () => {
            const node = { name: 'parent', children: [{ name: 'existing' }] };
            const child = { name: 'new' };
            xmlq.appendChild(node, child);
            expect(node).toEqualJson({ name: 'parent', children: [{ name: 'existing' }, { name: 'new' }] });
        });

        it("should create the children array if needed", () => {
            const node = { name: 'parent' };
            const child = { name: 'new' };
            xmlq.appendChild(node, child);
            expect(node).toEqualJson({ name: 'parent', children: [{ name: 'new' }] });
        });
    });

    describe("appendChildIfNotFound", () => {
        it("should append the child", () => {
            const node = { name: 'parent', children: [{ name: 'existing' }] };
            expect(xmlq.appendChildIfNotFound(node, 'new')).toEqualJson({ name: 'new', attributes: {}, children: [] });
            expect(node).toEqualJson({ name: 'parent', children: [{ name: 'existing' }, { name: 'new', attributes: {}, children: [] }] });
        });

        it("should not append the child", () => {
            const node = { name: 'parent', children: [{ name: 'existing' }] };
            expect(xmlq.appendChildIfNotFound(node, 'existing')).toEqualJson({ name: 'existing' });
            expect(node).toEqualJson({ name: 'parent', children: [{ name: 'existing' }] });
        });
    });

    describe("findChild", () => {
        it("should return the child", () => {
            const node = { name: 'parent', children: [{ name: 'A' }, { name: 'B' }] };
            expect(xmlq.findChild(node, 'A')).toBe(node.children[0]);
            expect(xmlq.findChild(node, 'B')).toBe(node.children[1]);
            expect(xmlq.findChild(node, 'C')).toBeUndefined();
        });
    });

    describe("getChildAttribute", () => {
        it("should return the child attribute", () => {
            const node = { name: 'parent', children: [
                { name: 'A' },
                { name: 'B', attributes: {} },
                { name: 'C', attributes: { foo: "FOO" } }
            ] };

            expect(xmlq.getChildAttribute(node, 'A', 'foo')).toBeUndefined();
            expect(xmlq.getChildAttribute(node, 'B', 'foo')).toBeUndefined();
            expect(xmlq.getChildAttribute(node, 'C', 'foo')).toBe('FOO');
        });
    });

    describe("hasChild", () => {
        it("should return true/false", () => {
            const node = { name: 'parent', children: [{ name: 'A' }, { name: 'B' }] };
            expect(xmlq.hasChild(node, 'A')).toBe(true);
            expect(xmlq.hasChild(node, 'B')).toBe(true);
            expect(xmlq.hasChild(node, 'C')).toBe(false);
        });
    });

    describe("insertAfter", () => {
        it("should insert the child after the node", () => {
            const node = { name: 'parent', children: [{ name: 'A' }, { name: 'B' }] };
            xmlq.insertAfter(node, { name: 'new' }, node.children[0]);
            expect(node.children).toEqualJson([{ name: 'A' }, { name: 'new' }, { name: 'B' }]);
        });
    });

    describe("insertBefore", () => {
        it("should insert the child before the node", () => {
            const node = { name: 'parent', children: [{ name: 'A' }, { name: 'B' }] };
            xmlq.insertBefore(node, { name: 'new' }, node.children[1]);
            expect(node.children).toEqualJson([{ name: 'A' }, { name: 'new' }, { name: 'B' }]);
        });
    });

    describe("insertInOrder", () => {
        it("should insert in the middle", () => {
            const node = { name: 'parent', children: [{ name: 'A' }, { name: 'C' }] };
            xmlq.insertInOrder(node, { name: 'B' }, ['A', 'B', 'C']);
            expect(node.children).toEqualJson([{ name: 'A' }, { name: 'B' }, { name: 'C' }]);
        });

        it("should insert at the beginning", () => {
            const node = { name: 'parent', children: [{ name: 'C' }] };
            xmlq.insertInOrder(node, { name: 'A' }, ['A', 'B', 'C']);
            expect(node.children).toEqualJson([{ name: 'A' }, { name: 'C' }]);
        });

        it("insert at the end", () => {
            const node = { name: 'parent', children: [{ name: 'A' }] };
            xmlq.insertInOrder(node, { name: 'C' }, ['A', 'B', 'C']);
            expect(node.children).toEqualJson([{ name: 'A' }, { name: 'C' }]);
        });

        it("append if node not expected in order", () => {
            const node = { name: 'parent', children: [{ name: 'A' }, { name: 'C' }] };
            xmlq.insertInOrder(node, { name: 'D' }, ['A', 'B', 'C']);
            expect(node.children).toEqualJson([{ name: 'A' }, { name: 'C' }, { name: 'D' }]);
        });
    });

    describe("isEmpty", () => {
        it("should return true/false", () => {
            const nodes = [
                { name: 'A' },
                { name: 'B', attributes: {}, children: [] },
                { name: 'C', attributes: { foo: 1 }, children: [] },
                { name: 'D', attributes: {}, children: [{}] },
                { name: 'E', attributes: { foo: 0 }, children: [{}] }
            ];

            expect(xmlq.isEmpty(nodes[0])).toBe(true);
            expect(xmlq.isEmpty(nodes[1])).toBe(true);
            expect(xmlq.isEmpty(nodes[2])).toBe(false);
            expect(xmlq.isEmpty(nodes[3])).toBe(false);
            expect(xmlq.isEmpty(nodes[4])).toBe(false);
        });
    });

    describe("removeChild", () => {
        it("should remove the children", () => {
            const node = { name: 'parent', children: [{ name: 'A' }, { name: 'B' }, { name: 'C' }] };
            xmlq.removeChild(node, node.children[1]);
            expect(node.children).toEqualJson([{ name: 'A' }, { name: 'C' }]);
            xmlq.removeChild(node, 'A');
            expect(node.children).toEqualJson([{ name: 'C' }]);
            xmlq.removeChild(node, 'foo');
            expect(node.children).toEqualJson([{ name: 'C' }]);
        });
    });

    describe("setAttributes", () => {
        it("should set/unset the attributes", () => {
            const node = { attributes: { foo: 1, bar: 1, baz: 1 } };
            xmlq.setAttributes(node, {
                foo: undefined,
                bar: null,
                goo: 1,
                gar: 1
            });
            expect(node.attributes).toEqualJson({
                baz: 1,
                goo: 1,
                gar: 1
            });
        });
    });

    describe("setChildAttributes", () => {
        it("should append the child with the attributes", () => {
            const node = { name: 'parent', children: [{ name: 'existing' }] };
            expect(xmlq.setChildAttributes(node, 'new', { foo: 1, bar: null })).toEqualJson({ name: 'new', attributes: { foo: 1 }, children: [] });
            expect(node.children).toEqualJson([{ name: 'existing' }, { name: 'new', attributes: { foo: 1 }, children: [] }]);
        });

        it("should not append the child but should set the attributes", () => {
            const node = { name: 'parent', children: [{ name: 'existing', attributes: { bar: 1 } }] };
            expect(xmlq.setChildAttributes(node, 'existing', { foo: 1, bar: null })).toEqualJson({ name: 'existing', attributes: { foo: 1 } });
            expect(node).toEqualJson({ name: 'parent', children: [{ name: 'existing', attributes: { foo: 1 } }] });
        });
    });

    describe("removeChildIfEmpty", () => {
        it("should remove the children", () => {
            const node = { name: 'parent', children: [{ name: 'A' }, { name: 'B' }, { name: 'C', attributes: { foo: 1 } }] };
            xmlq.removeChildIfEmpty(node, node.children[1]);
            expect(node.children).toEqualJson([{ name: 'A' }, { name: 'C', attributes: { foo: 1 } }]);
            xmlq.removeChildIfEmpty(node, 'A');
            expect(node.children).toEqualJson([{ name: 'C', attributes: { foo: 1 } }]);
            xmlq.removeChildIfEmpty(node, 'C');
            expect(node.children).toEqualJson([{ name: 'C', attributes: { foo: 1 } }]);
            xmlq.removeChildIfEmpty(node, node.children[0]);
            expect(node.children).toEqualJson([{ name: 'C', attributes: { foo: 1 } }]);
            xmlq.removeChildIfEmpty(node, 'foo');
            expect(node.children).toEqualJson([{ name: 'C', attributes: { foo: 1 } }]);
        });
    });
});
