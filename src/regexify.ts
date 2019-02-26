// https://stackoverflow.com/questions/3446170/escape-string-for-use-in-javascript-regex#6969486
function escapeRegExp(str: string): string {
    return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); // $& means the whole matched string
}

/**
 * Convert a pattern to a RegExp.
 * @param pattern - The pattern to convert.
 * @returns The regex.
 */
export function regexify(pattern: RegExp|string): RegExp {
    if (typeof pattern === 'string') {
        pattern = new RegExp(escapeRegExp(pattern), 'igm');
    }

    pattern.lastIndex = 0;

    return pattern;
}
