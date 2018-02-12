"use strict";

const proxyquire = require("proxyquire");

describe("ArgHandler", () => {
    let ArgHandler, argHandler, handlers, Style;

    beforeEach(() => {
        Style = class {}
        if (!Style.name) Style.name = "Style";

        ArgHandler = proxyquire("../../lib/ArgHandler", {
            '@noCallThru': true
        });

        handlers = {
            empty: jasmine.createSpy("empty").and.returnValue('empty'),
            nil: jasmine.createSpy("nil").and.returnValue("nil"),
            string: jasmine.createSpy("string").and.returnValue("string"),
            boolean: jasmine.createSpy("boolean").and.returnValue("boolean"),
            number: jasmine.createSpy("number").and.returnValue("number"),
            integer: jasmine.createSpy("integer").and.returnValue("integer"),
            function: jasmine.createSpy("function").and.returnValue("function"),
            array: jasmine.createSpy("array").and.returnValue("array"),
            date: jasmine.createSpy("date").and.returnValue("date"),
            object: jasmine.createSpy("object").and.returnValue("object"),
            Style: jasmine.createSpy("style").and.returnValue("Style"),
            '*': jasmine.createSpy("*").and.returnValue("*")
        };

        argHandler = new ArgHandler("METHOD")
            .case(handlers.empty)
            .case("nil", handlers.nil)
            .case("string", handlers.string)
            .case("boolean", handlers.boolean)
            .case("number", handlers.number)
            .case(["nil", "integer"], handlers.integer)
            .case("function", handlers.function)
            .case("array", handlers.array)
            .case("date", handlers.date)
            .case("object", handlers.object)
            .case("Style", handlers.Style)
            .case(["nil", 'nil', '*'], handlers['*']);
    });

    describe("handle", () => {
        it("should handle empty", () => {
            expect(argHandler.handle([])).toBe('empty');
            expect(handlers.empty).toHaveBeenCalledWith();
        });

        it("should handle nil", () => {
            expect(argHandler.handle([undefined])).toBe('nil');
            expect(handlers.nil).toHaveBeenCalledWith(undefined);
        });

        it("should handle string", () => {
            expect(argHandler.handle(["foo"])).toBe('string');
            expect(handlers.string).toHaveBeenCalledWith("foo");

            expect(argHandler.handle([""])).toBe('string');
            expect(handlers.string).toHaveBeenCalledWith("");
        });

        it("should handle boolean", () => {
            expect(argHandler.handle([true])).toBe('boolean');
            expect(handlers.boolean).toHaveBeenCalledWith(true);

            expect(argHandler.handle([false])).toBe('boolean');
            expect(handlers.boolean).toHaveBeenCalledWith(false);
        });

        it("should handle number", () => {
            expect(argHandler.handle([0])).toBe('number');
            expect(handlers.number).toHaveBeenCalledWith(0);

            expect(argHandler.handle([-5])).toBe('number');
            expect(handlers.number).toHaveBeenCalledWith(-5);

            expect(argHandler.handle([1.23])).toBe('number');
            expect(handlers.number).toHaveBeenCalledWith(1.23);
        });

        it("should handle integer", () => {
            expect(() => argHandler.handle([undefined, 1.5])).toThrow();

            expect(argHandler.handle([undefined, 3])).toBe('integer');
            expect(handlers.integer).toHaveBeenCalledWith(undefined, 3);

            expect(argHandler.handle([undefined, 0])).toBe('integer');
            expect(handlers.integer).toHaveBeenCalledWith(undefined, 0);

            expect(argHandler.handle([undefined, -5])).toBe('integer');
            expect(handlers.integer).toHaveBeenCalledWith(undefined, -5);
        });

        it("should handle function", () => {
            const func = () => {};
            expect(argHandler.handle([func])).toBe('function');
            expect(handlers.function).toHaveBeenCalledWith(func);
        });

        it("should handle array", () => {
            expect(argHandler.handle([[1, 2, 3]])).toBe('array');
            expect(handlers.array).toHaveBeenCalledWith([1, 2, 3]);
        });

        it("should handle date", () => {
            const date = new Date();
            expect(argHandler.handle([date])).toBe('date');
            expect(handlers.date).toHaveBeenCalledWith(date);
        });

        it("should handle object", () => {
            expect(argHandler.handle([{}])).toBe('object');
            expect(handlers.object).toHaveBeenCalledWith({});
        });

        it("should handle Styles", () => {
            const style = new Style();
            expect(argHandler.handle([style])).toBe('Style');
            expect(handlers.Style).toHaveBeenCalledWith(style);
        });

        it("should handle *", () => {
            expect(argHandler.handle([undefined, undefined, 1])).toBe('*');
            expect(handlers['*']).toHaveBeenCalledWith(undefined, undefined, 1);
        });
    });
});
