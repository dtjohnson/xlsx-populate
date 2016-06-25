"use strict";

var gulp = require("gulp");
var eslint = require("gulp-eslint");
var jasmine = require("gulp-jasmine");
var runSequence = require('run-sequence').use(gulp);

var TEST = "spec/**/*.spec.js";
var LIB = "lib/**/*.js";
var EXAMPLES = "examples/**/*.js";
var SRC = [LIB, TEST, EXAMPLES];

gulp.task("lint", function () {
    return gulp
        .src(SRC)
        .pipe(eslint())
        .pipe(eslint.format());
});

gulp.task("unit", function () {
    return gulp
        .src(TEST)
        .pipe(jasmine());
});

gulp.task("test", function (cb) {
    // Use run sequence to make sure lint and unit run in series. They both output to the
    // console to parallel execution leads to some funny output.
    runSequence("lint", "unit", cb);
});

gulp.task("watch", function () {
    gulp.watch(SRC, ["test"]);
});

gulp.task("default", function (cb) {
    runSequence("test", "watch", cb);
});
