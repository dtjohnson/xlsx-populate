"use strict";

var gulp = require("gulp");
var jasmine = require('gulp-jasmine');

var unitTests = "spec/**/*.spec.js";

gulp.task('unit', function () {
    return gulp.src(unitTests)
        .pipe(jasmine());
});

gulp.task('watch', function () {
    gulp.watch(unitTests, ['unit']);
});

gulp.task('default', ['unit', 'watch']);
