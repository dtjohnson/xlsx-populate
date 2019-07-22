import { OverloadHandler } from '../../src/OverloadHandler';

describe('OverloadHandler', () => {
    let overloadHandler: OverloadHandler;
    let handlers: { [name: string]: any };

    class SomeClass {}

    beforeEach(() => {
        handlers = {
            empty: jasmine.createSpy('empty').and.returnValue('empty'),
            nil: jasmine.createSpy('nil').and.returnValue('nil'),
            string: jasmine.createSpy('string').and.returnValue('string'),
            boolean: jasmine.createSpy('boolean').and.returnValue('boolean'),
            number: jasmine.createSpy('number').and.returnValue('number'),
            function: jasmine.createSpy('function').and.returnValue('function'),
            array: jasmine.createSpy('array').and.returnValue('array'),
            date: jasmine.createSpy('date').and.returnValue('date'),
            object: jasmine.createSpy('object').and.returnValue('object'),
            SomeClass: jasmine.createSpy('SomeClass').and.returnValue('SomeClass'),
            any: jasmine.createSpy('any').and.returnValue('any'),
        };

        overloadHandler = new OverloadHandler('METHOD')
            .case(handlers.empty)
            .case(undefined, handlers.nil)
            .case('string', handlers.string)
            .case('boolean', handlers.boolean)
            .case('number', handlers.number)
            .case(Function, handlers.function)
            .case(Array, handlers.array)
            .case(Date, handlers.date)
            .case(SomeClass, handlers.SomeClass)
            .case(Object, handlers.object)
            .case(undefined, undefined, 'any', handlers.any);
    });

    describe('handle', () => {
        it('should handle empty', () => {
            expect(overloadHandler.handle([])).toBe('empty');
            expect(handlers.empty).toHaveBeenCalledWith();
        });

        it('should handle nil', () => {
            expect(overloadHandler.handle([ undefined ])).toBe('nil');
            expect(handlers.nil).toHaveBeenCalledWith(undefined);
        });

        it('should handle string', () => {
            expect(overloadHandler.handle([ 'foo' ])).toBe('string');
            expect(handlers.string).toHaveBeenCalledWith('foo');

            expect(overloadHandler.handle([ '' ])).toBe('string');
            expect(handlers.string).toHaveBeenCalledWith('');
        });

        it('should handle boolean', () => {
            expect(overloadHandler.handle([ true ])).toBe('boolean');
            expect(handlers.boolean).toHaveBeenCalledWith(true);

            expect(overloadHandler.handle([ false ])).toBe('boolean');
            expect(handlers.boolean).toHaveBeenCalledWith(false);
        });

        it('should handle number', () => {
            expect(overloadHandler.handle([ 0 ])).toBe('number');
            expect(handlers.number).toHaveBeenCalledWith(0);

            expect(overloadHandler.handle([ -5 ])).toBe('number');
            expect(handlers.number).toHaveBeenCalledWith(-5);

            expect(overloadHandler.handle([ 1.23 ])).toBe('number');
            expect(handlers.number).toHaveBeenCalledWith(1.23);
        });

        it('should handle function', () => {
            const func = () => {};
            expect(overloadHandler.handle([ func ])).toBe('function');
            expect(handlers.function).toHaveBeenCalledWith(func);
        });

        it('should handle array', () => {
            expect(overloadHandler.handle([ [ 1, 2, 3 ] ])).toBe('array');
            expect(handlers.array).toHaveBeenCalledWith([ 1, 2, 3 ]);
        });

        it('should handle date', () => {
            const date = new Date();
            expect(overloadHandler.handle([ date ])).toBe('date');
            expect(handlers.date).toHaveBeenCalledWith(date);
        });

        it('should handle object', () => {
            expect(overloadHandler.handle([ {} ])).toBe('object');
            expect(handlers.object).toHaveBeenCalledWith({});
        });

        it('should handle SomeClass', () => {
            const someInstance = new SomeClass();
            expect(overloadHandler.handle([ someInstance ])).toBe('SomeClass');
            expect(handlers.SomeClass).toHaveBeenCalledWith(someInstance);
        });

        it('should handle any', () => {
            expect(overloadHandler.handle([ undefined, undefined, 1 ])).toBe('any');
            expect(handlers.any).toHaveBeenCalledWith(undefined, undefined, 1);
        });
    });
});
