"use strict";

const XlsxPoplate = require('../../lib/XlsxPopulate');
const RichText = require('../../lib/RichText');
const RichTextFragment = require('../../lib/RichTextFragment');

describe("RichText", () => {
    let cell, workbook, cell2, cell3;

    beforeEach(done => {
        XlsxPoplate.fromBlankAsync()
            .then(wb => {
                workbook = wb;
                cell = workbook.sheet(0).cell(1, 1);
                cell2 = workbook.sheet(0).cell(1, 2);
                cell3 = workbook.sheet(0).cell(1, 3);
                done();
            });
    });

    it('global export', () => {
        expect(RichText === XlsxPoplate.RichText).toBe(true);
    });

    describe("add/get", () => {
        it("should add and get normal text", () => {
            const rt = new RichText();
            cell.value(rt);
            expect(cell.value() instanceof RichText).toBe(true);
            rt.add('hello');
            rt.add('hello2');
            expect(rt.length).toBe(2);
            expect(rt.get(0).value()).toBe('hello');
            expect(rt.get(1).value()).toBe('hello2');
        });

        it("should transfer line separator to \r\n", () => {
            const rt = new RichText();
            rt.add('hello\n');
            rt.add('hel\r\nlo2');
            rt.add('hel\rlo2');
            cell.value(rt);
            expect(rt.get(0).value()).toBe('hello\r\n');
            expect(rt.get(1).value()).toBe('hel\r\nlo2');
            expect(rt.get(2).value()).toBe('hel\r\nlo2');
        });

        it("should set wrapText to true", () => {
            const rt = new RichText();
            rt.add('hello\n');
            cell.value(rt);
            expect(cell.style('wrapText')).toBe(true);
        });

        it("should set style", () => {
            const rt = new RichText();
            rt.add('hello\n', { bold: true });
            cell.value(rt);
            expect(rt.get(0).style('bold')).toBe(true);
        });
    });

    it("should clear the rich text", () => {
        const rt = new RichText();
        rt.add('hello');
        rt.clear();
        expect(rt.text()).toBe('');
    });

    it("should get concatenated text", () => {
        const rt = new RichText();
        rt.add('hello')
            .add(' I', { fontColor: 'FF0000FF' })
            .add("'m \n ")
            .add('lester');

        expect(rt.text()).toBe("hello I'm \r\n lester");
    });

    describe("change related cell", () => {
        it("should set wrapText to true in the new cell", () => {
            const rt = new RichText();
            rt.add('hello\n');
            cell.value(rt);
            expect(cell.style('wrapText')).toBe(true);
        });
    });

    describe('Cell.value', () => {
        it('should assign a deep copy of rich text instance', () => {
            const rt = new RichText();
            rt.add('string\n');
            cell.value(rt);
            cell2.value(rt);
            const value1 = cell.value(), value2 = cell2.value();
            expect(value1).not.toBe(rt);
            expect(value2).not.toBe(rt);
            expect(value1).not.toBe(value2);

            value1.add('test');
            expect(cell.value().text()).toBe('string\r\ntest');
            expect(cell2.value().text()).toBe('string\r\n');
        });

        it('should get instance with cell reference', () => {
            const rt = new RichText();
            rt.add('string');
            cell.value(rt);
            expect(cell.value().cell).toBe(cell);
        });

        it('should re-assign cell reference', () => {
            const rt = new RichText();
            rt.add('string');
            cell.value(rt);
            expect(cell.value().cell).toBe(cell);
            const value = cell.value();
            cell2.value(value);
            expect(cell2.value().cell).toBe(cell2);
        });
    });

    describe('Sheet.range', () => {
        it('should set range of rich texts', () => {
            const rt = new RichText();
            rt.add('string');
            workbook.sheet(0).range('A1:C3').value(rt);
            expect(cell.value().cell).toBe(cell);
            expect(cell.value().text()).toBe('string');
            expect(cell.value()).not.toBe(rt);
            expect(cell.value()).not.toBe(cell2.value());
        });
    });

    describe('styles', () => {
        let fontNode, fragment;

        beforeEach(() => {
            fragment = new RichTextFragment('text');
            fontNode = fragment._fontNode;
        });

        describe("bold", () => {
            it("should get/set bold", () => {
                expect(fragment.style("bold")).toBe(false);
                fragment.style("bold", true);
                expect(fragment.style("bold")).toBe(true);
                expect(fontNode.children).toEqualJson([{ name: "b", attributes: {}, children: [] }]);
                fragment.style("bold", false);
                expect(fragment.style("bold")).toBe(false);
                expect(fontNode.children).toEqualJson([]);
            });
        });

        describe("italic", () => {
            it("should get/set italic", () => {
                expect(fragment.style("italic")).toBe(false);
                fragment.style("italic", true);
                expect(fragment.style("italic")).toBe(true);
                expect(fontNode.children).toEqualJson([{ name: "i", attributes: {}, children: [] }]);
                fragment.style("italic", false);
                expect(fragment.style("italic")).toBe(false);
                expect(fontNode.children).toEqualJson([]);
            });
        });

        describe("underline", () => {
            it("should get/set underline", () => {
                expect(fragment.style("underline")).toBe(false);
                fragment.style("underline", true);
                expect(fragment.style("underline")).toBe(true);
                expect(fontNode.children).toEqualJson([{ name: "u", attributes: {}, children: [] }]);
                fragment.style("underline", "double");
                expect(fragment.style("underline")).toBe("double");
                expect(fontNode.children).toEqualJson([{ name: "u", attributes: { val: "double" }, children: [] }]);
                fragment.style("underline", true);
                expect(fragment.style("underline")).toBe(true);
                expect(fontNode.children).toEqualJson([{ name: "u", attributes: {}, children: [] }]);
                fragment.style("underline", false);
                expect(fragment.style("underline")).toBe(false);
                expect(fontNode.children).toEqualJson([]);
            });
        });

        describe("strikethrough", () => {
            it("should get/set strikethrough", () => {
                expect(fragment.style("strikethrough")).toBe(false);
                fragment.style("strikethrough", true);
                expect(fragment.style("strikethrough")).toBe(true);
                expect(fontNode.children).toEqualJson([{ name: 'strike', attributes: {}, children: [] }]);
                fragment.style("strikethrough", false);
                expect(fragment.style("strikethrough")).toBe(false);
                expect(fontNode.children).toEqualJson([]);
            });
        });

        describe("subscript", () => {
            it("should get/set subscript", () => {
                expect(fragment.style("subscript")).toBe(false);
                fragment.style("subscript", true);
                expect(fragment.style("subscript")).toBe(true);
                expect(fontNode.children).toEqualJson([{
                    name: "vertAlign",
                    attributes: { val: "subscript" },
                    children: []
                }]);
                fragment.style("subscript", false);
                expect(fragment.style("subscript")).toBe(false);
                expect(fontNode.children).toEqualJson([]);
            });
        });

        describe("superscript", () => {
            it("should get/set superscript", () => {
                expect(fragment.style("superscript")).toBe(false);
                fragment.style("superscript", true);
                expect(fragment.style("superscript")).toBe(true);
                expect(fontNode.children).toEqualJson([{
                    name: "vertAlign",
                    attributes: { val: "superscript" },
                    children: []
                }]);
                fragment.style("superscript", false);
                expect(fragment.style("superscript")).toBe(false);
                expect(fontNode.children).toEqualJson([]);
            });
        });

        describe("fontSize", () => {
            it("should get/set fontSize", () => {
                expect(fragment.style("fontSize")).toBe(undefined);
                fragment.style("fontSize", 17);
                expect(fragment.style("fontSize")).toBe(17);
                expect(fontNode.children).toEqualJson([{ name: 'sz', attributes: { val: 17 }, children: [] }]);
                fragment.style("fontSize", undefined);
                expect(fragment.style("fontSize")).toBe(undefined);
                expect(fontNode.children).toEqualJson([]);
            });
        });

        describe("fontFamily", () => {
            it("should get/set fontFamily", () => {
                expect(fragment.style("fontFamily")).toBe(undefined);
                fragment.style("fontFamily", "Comic Sans MS");
                expect(fragment.style("fontFamily")).toBe("Comic Sans MS");
                expect(fontNode.children).toEqualJson([{
                    name: 'rFont',
                    attributes: { val: "Comic Sans MS" },
                    children: []
                }]);
                fragment.style("fontFamily", undefined);
                expect(fragment.style("fontFamily")).toBe(undefined);
                expect(fontNode.children).toEqualJson([]);
            });
        });

        describe("fontGenericFamily", () => {
            it("should get/set fontGenericFamily", () => {
                expect(fragment.style("fontGenericFamily")).toBe(undefined);
                fragment.style("fontGenericFamily", 1);
                expect(fragment.style("fontGenericFamily")).toBe(1);
                expect(fontNode.children).toEqualJson([{ name: 'family', attributes: { val: 1 }, children: [] }]);
                fragment.style("fontGenericFamily", undefined);
                expect(fragment.style("fontGenericFamily")).toBe(undefined);
                expect(fontNode.children).toEqualJson([]);
            });
        });

        describe("fontScheme", () => {
            it("should get/set fontScheme", () => {
                expect(fragment.style("fontScheme")).toBe(undefined);
                fragment.style("fontScheme", 'minor');
                expect(fragment.style("fontScheme")).toBe('minor');
                expect(fontNode.children).toEqualJson([{ name: 'scheme', attributes: { val: 'minor' }, children: [] }]);
                fragment.style("fontScheme", undefined);
                expect(fragment.style("fontScheme")).toBe(undefined);
                expect(fontNode.children).toEqualJson([]);
            });
        });

        describe("fontColor", () => {
            it("should get/set fontColor", () => {
                expect(fragment.style("fontColor")).toBe(undefined);

                fragment.style("fontColor", "ff0000");
                expect(fragment.style("fontColor")).toEqualJson({ rgb: "FF0000" });
                expect(fontNode.children).toEqualJson([{ name: 'color', attributes: { rgb: "FF0000" }, children: [] }]);

                fragment.style("fontColor", 5);
                expect(fragment.style("fontColor")).toEqualJson({ theme: 5 });
                expect(fontNode.children).toEqualJson([{ name: 'color', attributes: { theme: 5 }, children: [] }]);

                fragment.style("fontColor", { theme: 3, tint: -0.2 });
                expect(fragment.style("fontColor")).toEqualJson({ theme: 3, tint: -0.2 });
                expect(fontNode.children).toEqualJson([{
                    name: 'color',
                    attributes: { theme: 3, tint: -0.2 },
                    children: []
                }]);

                fragment.style("fontColor", undefined);
                expect(fragment.style("fontColor")).toBe(undefined);
                expect(fontNode.children).toEqualJson([]);

                fontNode.children = [{ name: 'color', attributes: { indexed: 7 }, children: [] }];
                expect(fragment.style("fontColor")).toEqualJson({ rgb: "00FFFF" });
            });
        });
    });

    it('should preserve leading space', () => {
        const rt = new RichText();
        rt.add(' hello').add(' I');
        expect(rt.text()).toBe(" hello I");
        expect(rt.get(0)._valueNode.attributes['xml:space']).toBe('preserve');
        expect(rt.get(1)._valueNode.attributes['xml:space']).toBe('preserve');
    });

    xit('it should save unsupported node', done => {
        XlsxPoplate.fromFileAsync('./test/files/issue-230.xlsx')
            .then(wb => {
                expect(wb.sheet(0).cell('A1').value()._remainingNodes.length).toBe(5);
                done();
            });
    });
});
