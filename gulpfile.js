"use strict";

var gulp = require("gulp");
var cached = require("gulp-cached");
var eslint = require("gulp-eslint");
var jasmine = require("gulp-jasmine");

var TEST = "spec/**/*.spec.js";
var LIB = "lib/**/*.js";
var EXAMPLES = "examples/**/*.js";
var SRC = [LIB, TEST, EXAMPLES];

gulp.task("lint", function () {
    return gulp
        .src(SRC)
        .pipe(cached("lint"))
        .pipe(eslint())
        .pipe(eslint.format())
        .pipe(eslint.failAfterError())
        ;
});

gulp.task("unit", ["lint"], function () {
    return gulp
        .src(TEST)
        .pipe(cached("unit"))
        .pipe(jasmine())
        ;
});

gulp.task("watch", function () {
    gulp.watch(SRC, ["unit"]);
});

gulp.task("default", ["watch"]);
