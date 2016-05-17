"use strict";

var gulp = require("gulp");
var jasmine = require('gulp-jasmine');

var unitTests = "spec/**/*.spec.js";
var jsFiles = "lib/**/*.js";

gulp.task('unit', function () {
    return gulp.src(unitTests)
        .pipe(jasmine());
});

gulp.task('watch', function () {
    gulp.watch([jsFiles, unitTests], ['unit']);
});

gulp.task('default', ['unit', 'watch']);
