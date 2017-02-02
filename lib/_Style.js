"use strict";

const _ArgParser = require("./_ArgParser");
const debug = require("./debug")("_Style");
const jq = require("./jq");
const _ = require("lodash");

class Style {
    constructor(styleSheet, id, xfNode, fontNode, fillNode, borderNode) {
        this._styleSheet = styleSheet;
        this._id = id;
        this._xfNode = xfNode;
        this._fontNode = fontNode;
        this._fillNode = fillNode;
        this._borderNode = borderNode;
    }

    style() {
        debug("style(%o)", arguments);
        return new _ArgParser("_Style.style")
            .case(String, name => {
                const methodName = `__${name}`;
                if (!this[methodName]) throw new Error(`_Style.style: '${name}' is not a valid style`);
                return this[methodName]();
            })
            .case([String, undefined], (name, value) => {
                const methodName = `__${name}`;
                if (!this[methodName]) throw new Error(`_Style.style: '${name}' is not a valid style`);
                this[methodName](value);
                return this;
            })
            .parse(arguments);
    }

    __bold(bold) {
        debug("__bold(%o)", arguments);
        if (arguments.length === 0) return _.has(this._fontNode, "b[0]");
        if (bold) jq.setIfNeeded(this._fontNode, "b[0]", {});
        else delete this._fontNode.b;
    }

    __italic(italic) {
        debug("__italic(%o)", arguments);
        if (arguments.length === 0) return jq.has(this._fontNode, "i[0]");
        if (italic) jq.setIfNeeded(this._fontNode, "i[0]", {});
        else delete this._fontNode.i;
    }

    __underline(underline) {
        debug("__underline(%o)", arguments);
        if (arguments.length === 0) {
            const val = jq.get(this._fontNode, "u[0].$.val");
            if (val) return val;
            return jq.has(this._fontNode, "u[0]");
        }

        if (typeof underline === "string") jq.set(this._fontNode, "u[0].$.val", underline);
        else if (underline) jq.setIfNeeded(this._fontNode, "strike[0]", {});
        else delete this._fontNode.u;
    }

    __strikethrough(strikethrough) {
        debug("__strikethrough(%o)", arguments);
        if (arguments.length === 0) return jq.has(this._fontNode, "strike[0]");
        if (strikethrough) jq.setIfNeeded(this._fontNode, "strike[0]", {});
        else delete this._fontNode.strike;
    }

    __fontVerticalAlignment(alignment) {
        debug("__fontVerticalAlignment(%o)", arguments);
        if (arguments.length === 0) return jq.get(this._fontNode, "vertAlign[0].$.val");
        jq.set(this._fontNode, "vertAlign[0].$.val", alignment);
        if (jq.isEmpty(this._fontNode, "vertAlign[0].$")) delete this._fontNode.vertAlign;
    }

    __superscript(superscript) {
        debug("__superscript(%o)", arguments);
        if (arguments.length === 0) return this.__fontVerticalAlignment() === "superscript";
        return this.__fontVerticalAlignment(superscript ? "superscript" : null);
    }

    __subscript(subscript) {
        debug("__subscript(%o)", arguments);
        if (arguments.length === 0) return this.__fontVerticalAlignment() === "subscript";
        return this.__fontVerticalAlignment(subscript ? "subscript" : null);
    }

    __fontSize(size) {
        debug("__fontSize(%o)", arguments);
        if (arguments.length === 0) return jq.get(this._fontNode, "sz[0].$.val");
        jq.set(this._fontNode, "sz[0].$.val", size);
        if (jq.isEmpty(this._fontNode, "sz[0].$")) delete this._fontNode.sz;
    }

    __fontFamily(family) {
        debug("__fontFamily(%o)", arguments);
        if (arguments.length === 0) return jq.get(this._fontNode, "name[0].$.val");
        jq.set(this._fontNode, "name[0].$.val", family);
        if (jq.isEmpty(this._fontNode, "name[0].$")) delete this._fontNode.name;
    }

    // TODO: # prefix?, rgb(x, x, x)?
    __fontColor(color) {
        debug("__fontColor(%o)", arguments);
        if (arguments.length === 0) {
            return jq.apply(this._fontNode, "color[0].$", $ => $ && $.rgb || $.indexed);
        }

        let rgb, indexed;
        if (typeof color === "string") rgb = color;
        else if (color >= 0) indexed = color;

        jq.set(this._fontNode, {
            "color[0].$.rgb": rgb,
            "color[0].$.indexed": indexed
        });

        if (jq.isEmpty(this._fontNode, "color[0].$")) delete this._fontNode.color;
    }

    __horizontalAlignment(alignment) {
        debug("__horizontalAlignment(%o)", arguments);
        if (arguments.length === 0) return jq.get(this._xfNode, "alignment[0].$.horizontal");
        jq.set(this._xfNode, "alignment[0].$.horizontal", alignment);
        if (jq.isEmpty(this._xfNode, "alignment[0].$")) delete this._xfNode.alignment;
    }

    __verticalAlignment(alignment) {
        debug("__verticalAlignment(%o)", arguments);
        if (arguments.length === 0) return jq.get(this._xfNode, "alignment[0].$.vertical");
        jq.set(this._xfNode, "alignment[0].$.vertical", alignment);
        if (jq.isEmpty(this._xfNode, "alignment[0].$")) delete this._xfNode.alignment;
    }

    __wrappedText(wrappedText) {
        debug("__wrappedText(%o)", arguments);
        if (arguments.length === 0) return jq.get(this._xfNode, "alignment[0].$.wrapText");
        jq.set(this._xfNode, "alignment[0].$.wrapText", wrappedText);
        if (jq.isEmpty(this._xfNode, "alignment[0].$")) delete this._xfNode.alignment;
    }

    __indent(indent) {
        debug("__indent(%o)", arguments);
        if (arguments.length === 0) return jq.get(this._xfNode, "alignment[0].$.indent");
        jq.set(this._xfNode, "alignment[0].$.indent", indent);
        if (jq.isEmpty(this._xfNode, "alignment[0].$")) delete this._xfNode.alignment;
    }

    // TODO: Negative values?
    __textRotation(textRotation) {
        debug("__textRotation(%o)", arguments);
        if (arguments.length === 0) return jq.get(this._xfNode, "alignment[0].$.textRotation");
        jq.set(this._xfNode, "alignment[0].$.textRotation", textRotation);
        if (jq.isEmpty(this._xfNode, "alignment[0].$")) delete this._xfNode.alignment;
    }

    __angleTextCounterclockwise(value) {
        debug("__angleTextCounterclockwise(%o)", arguments);
        if (arguments.length === 0) return this.__textRotation() === 45;
        return this.__textRotation(value ? 45 : null);
    }

    __angleTextClockwise(value) {
        debug("__angleTextClockwise(%o)", arguments);
        if (arguments.length === 0) return this.__textRotation() === 135;
        return this.__textRotation(value ? 135 : null);
    }

    __verticalText(value) {
        debug("__verticalText(%o)", arguments);
        if (arguments.length === 0) return this.__textRotation() === 255;
        return this.__textRotation(value ? 255 : null);
    }

    __rotateTextUp(value) {
        debug("__rotateTextUp(%o)", arguments);
        if (arguments.length === 0) return this.__textRotation() === 90;
        return this.__textRotation(value ? 90 : null);
    }

    __rotateTextDown(value) {
        debug("__rotateTextDown(%o)", arguments);
        if (arguments.length === 0) return this.__textRotation() === 180;
        return this.__textRotation(value ? 180 : null);
    }

    __numberFormat(formatCode) {
        debug("__numberFormat(%o)", arguments);
        if (arguments.length === 0) {
            const numFmtId = jq.get(this._xfNode, '$.numFmtId') || 0;
            return this._styleSheet.getNumberFormatCode(numFmtId);
        }

        jq.set(this._xfNode, '$.numFmtId', this._styleSheet.getNumberFormatId(formatCode));
    }

    // TODO: Consider various border options. Should we merge like CSS borders?

    _borderStyle(side, args) {
        debug("_borderStyle(%o)", arguments);
        if (args.length === 0) return jq.get(this._borderNode, `${side}[0].$.style`);
        jq.set(this._borderNode, `${side}[0].$.style`, args[0]);
        if (jq.isEmpty(this._borderNode, `${side}[0].$`)) delete this._borderNode[side];
    }

    __topBorderStyle() {
        debug("__topBorderStyle(%o)", arguments);
        return this._borderStyle("top", arguments);
    }

    __bottomBorderStyle() {
        debug("__bottomBorderStyle(%o)", arguments);
        return this._borderStyle("bottom", arguments);
    }

    __leftBorderStyle() {
        debug("leftBorderStyle(%o)", arguments);
        return this._borderStyle("left", arguments);
    }

    __rightBorderStyle() {
        debug("rightBorderStyle(%o)", arguments);
        return this._borderStyle("right", arguments);
    }
}

module.exports = Style;

/*

 Fill:
 style with applyFill="1" -> fill with:
 <patternFill patternType="solid">
 <fgColor rgb="FFFFFF00"/>
 <bgColor indexed="64"/>
 </patternFill>

 Font color x:
 style -> font with <color rgb="x"/> indexed, theme, or rgb

 Border:
 style with applyBorder="1" -> border:
 <border>
 <left /> style="thin|medium|thick|double"
 child <color> indexed = 64 is black (not sure if necessary)
 ...
 </border>

 Number Format:
 style -> numFmtId
 Create numFmt if not standard: http://polymathprogrammer.com/2011/02/15/built-in-styles-for-excel-open-xml/
 */
