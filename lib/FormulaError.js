"use strict";

const _ = require("lodash");

class FormulaError {
    constructor(error) {
        this._error = error;
    }

    error() {
        return this._error;
    }
}

FormulaError.DIV0 = new FormulaError("#DIV/0!");
FormulaError.NA = new FormulaError("#N/A");
FormulaError.NAME = new FormulaError("#NAME?");
FormulaError.NULL = new FormulaError("#NULL!");
FormulaError.NUM = new FormulaError("#NUM!");
FormulaError.REF = new FormulaError("#REF!");
FormulaError.VALUE = new FormulaError("#VALUE!");

FormulaError.getError = error => {
    return _.find(FormulaError, value => {
        return value instanceof FormulaError && value.error() === error;
    });
};

module.exports = FormulaError;
