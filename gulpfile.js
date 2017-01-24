"use strict";

const gulp = require('gulp');

const browserify = require('browserify');
const babelify = require('babelify');
const source = require('vinyl-source-stream');
const buffer = require('vinyl-buffer');
const uglify = require('gulp-uglify');
const sourcemaps = require('gulp-sourcemaps');

const BROWSERIFY_STANDALONE_NAME = "XLSXPopulate";
const BABEL_PRESETS = ["es2015"];
const PATHS = {
    browserify: {
        source: "./lib/browser.js",
        base: "./browser",
        bundle: "xlsx-populate.js",
        sourceMap: "./"
    }
};

gulp.task('build', () => {
    return browserify({
        entries: PATHS.browserify.source,
        debug: true,
        standalone: BROWSERIFY_STANDALONE_NAME
    })
        .transform("babelify", { presets: BABEL_PRESETS })
        .bundle()
        .pipe(source(PATHS.browserify.base))
        .pipe(buffer())
        .pipe(sourcemaps.init({ loadMaps: true }))
        .pipe(uglify())
        .pipe(sourcemaps.write(PATHS.browserify.sourceMap))
        .pipe(gulp.dest(PATHS.browserify.bundle));
});

gulp.task('watch', ['build'], function () {
    gulp.watch('./lib/**/*.js', ['build']);
});

gulp.task('default', ['watch']);

//
// var gulp = require("gulp");
// var eslint = require("gulp-eslint");
// var jasmine = require("gulp-jasmine");
// var runSequence = require('run-sequence').use(gulp);
//
// var sourcemaps = require('gulp-sourcemaps');
// var source = require('vinyl-source-stream');
// var buffer = require('vinyl-buffer');
// var browserify = require('browserify');
// var watchify = require('watchify');
// var babel = require('babelify');
//
//
// function compile(watch) {
//     var bundler = watchify(browserify('./lib/browser.js', { debug: true, standalone: "XLSXPopulate" }).transform(babel, { presets: ["es2015"]}));
//
//     function rebundle() {
//         bundler.bundle()
//             .on('error', function(err) { console.error(err); this.emit('end'); })
//             .pipe(source('xlsx-populate.js'))
//             .pipe(buffer())
//             .pipe(sourcemaps.init({ loadMaps: true }))
//             .pipe(sourcemaps.write('./'))
//             .pipe(gulp.dest('./browser'));
//     }
//
//     if (watch) {
//         bundler.on('update', function() {
//             console.log('-> bundling...');
//             rebundle();
//         });
//     }
//
//     rebundle();
// }
//
// function watch() {
//     return compile(true);
// };
//
// gulp.task('build', function() { return compile(); });
// gulp.task('watch', function() { return watch(); });
//
// gulp.task('default', ['watch']);

// var TEST = "spec/**/*.spec.js";
// var LIB = "lib/**/*.js";
// var EXAMPLES = "examples/**/*.js";
// var SRC = [LIB, TEST, EXAMPLES];
//
//
// gulp.task("lint", function () {
//     return gulp
//         .src(SRC)
//         .pipe(eslint())
//         .pipe(eslint.format());
// });
//
// gulp.task("unit", function () {
//     return gulp
//         .src(TEST)
//         .pipe(jasmine({
//             includeStackTrace: false,
//             errorOnFail: false
//         }));
// });
//
// gulp.task("test", function (cb) {
//     // Use run sequence to make sure lint and unit run in series. They both output to the
//     // console so parallel execution would lead to some funny output.
//     runSequence("unit", cb);//"lint"
// });
//
// gulp.task("watch", function () {
//     gulp.watch(SRC, ["test"]);
// });
//
// gulp.task("default", function (cb) {
//     runSequence("test", "watch", cb);
// });
