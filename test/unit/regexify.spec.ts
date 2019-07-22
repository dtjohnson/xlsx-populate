import { regexify } from '../../src/regexify';

describe('regexify', () => {
    it('should return the regex with lastIndex reset', () => {
        const regexp = /.+/;
        regexp.lastIndex = 5;

        const actual = regexify(regexp);
        expect(actual).toBe(regexp);
        expect(actual.lastIndex).toBe(0);
    });

    it('should convert a string to a regexp', () => {
        expect(regexify('search.[?')).toEqual(/search\.\[\?/gim);
    });
});
