/**
 * Remove any code points not allowed in the XML specification.
 * https://www.w3.org/TR/2008/REC-xml-20081126/#charsets
 * @param {string} string - the string to sanitize
 * @ignore
 */

const INVALID_XML_REGEX =  /([^\x09\x0A\x0D\x20-\uD7FF\uE000-\uFFFC\u{10000}-\u{10FFFF}])/ug;

function xmlSantitize(string) {
  return string.replace(INVALID_XML_REGEX, '');
}

module.exports = xmlSantitize;
