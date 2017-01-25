"use strict";

const gulp = require('gulp');

const browserify = require('browserify');
const babelify = require('babelify');
const source = require('vinyl-source-stream');
const buffer = require('vinyl-buffer');
const uglify = require('gulp-uglify');
const sourcemaps = require('gulp-sourcemaps');
const eslint = require("gulp-eslint");
const jasmine = require("gulp-jasmine");
const runSequence = require('run-sequence').use(gulp);

const BROWSERIFY_STANDALONE_NAME = "XLSXPopulate";
const BABEL_PRESETS = ["es2015"];
const PATHS = {
    lib: "./lib/**/*.js",
    spec: "./spec/**/*.js",
    examples: "./examples/**/*.js",
    browserify: {
        source: "./lib/Workbook.js",
        base: "./browser",
        bundle: "xlsx-populate.js",
        sourceMap: "./"
    }
};

PATHS.lint = [PATHS.lib];
PATHS.testSources = [PATHS.lib, PATHS.spec];

gulp.task('build', () => {
    return browserify({
        entries: PATHS.browserify.source,
        debug: true,
        standalone: BROWSERIFY_STANDALONE_NAME
    })
        .transform("babelify", { presets: BABEL_PRESETS })
        .transform("brfs")
        .bundle()
        .pipe(source(PATHS.browserify.bundle))
        .pipe(buffer())
        .pipe(sourcemaps.init({ loadMaps: true }))
        .pipe(uglify())
        .pipe(sourcemaps.write(PATHS.browserify.sourceMap))
        .pipe(gulp.dest(PATHS.browserify.base));
});

gulp.task("lint", () => {
    return gulp
        .src(PATHS.lint)
        .pipe(eslint())
        .pipe(eslint.format());
});

gulp.task("unit", () => {
    return gulp
        .src(PATHS.spec)
        .pipe(jasmine({
            includeStackTrace: false,
            errorOnFail: false
        }));
});


gulp.task("test", cb => {
    // Use run sequence to make sure lint and unit run in series. They both output to the
    // console so parallel execution would lead to some funny output.
    runSequence("unit", cb);//"lint"
});

gulp.task('watch', ['build'], () => {
    gulp.watch(PATHS.lib, ['build']);
    gulp.watch(PATHS.testSources, ["test"]);
});

gulp.task("default", cb => {
    runSequence(["build", "test"], "watch", cb);
});
