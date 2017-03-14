"use strict";

/* eslint global-require: "off" */

const gulp = require('gulp');
const browserify = require('browserify');
const babelify = require('babelify');
const source = require('vinyl-source-stream');
const buffer = require('vinyl-buffer');
const uglify = require('gulp-uglify');
const sourcemaps = require('gulp-sourcemaps');
const eslint = require("gulp-eslint");
const runSequence = require('run-sequence').use(gulp);
const jsdoc2md = require("jsdoc-to-markdown");
const toc = require('markdown-toc');
const Promise = require("bluebird");
const fs = Promise.promisifyAll(require("fs"));
const karma = require('karma');
const Jasmine = require("jasmine");

const BROWSERIFY_STANDALONE_NAME = "XlsxPopulate";
const BABEL_CONFIG = { presets: ["es2015"] };
const PATHS = {
    lib: "./lib/**/*.js",
    unit: "./test/unit/**/*.js",
    karma: ["./test/helpers/**/*.js", "./test/unit/**/*.spec.js"], // Helpers need to go first
    examples: "./examples/**/*.js",
    browserify: {
        source: "./lib/XlsxPopulate.js",
        base: "./browser",
        bundle: "xlsx-populate.js",
        sourceMap: "./"
    },
    readme: {
        template: "./docs/template.md",
        build: "./README.md"
    },
    blank: {
        workbook: "./blank/blank.xlsx",
        template: "./blank/template.js",
        build: "./lib/blank.js"
    },
    jasmineConfigs: {
        unit: "./test/unit/jasmine.json",
        e2eGenerate: "./test/e2e-generate/jasmine.json",
        e2eParse: "./test/e2e-parse/jasmine.json"
    }
};

PATHS.lint = [PATHS.lib];
PATHS.unitTestSources = [PATHS.lib, PATHS.unit];

// Function to clear the require cache as running unit tests mess up later tests.
const clearRequireCache = () => {
    for (const moduleId in require.cache) {
        if (require.cache.hasOwnProperty(moduleId)) {
            delete require.cache[moduleId];
        }
    }
};

const runKarma = (files, cb) => {
    process.chdir(__dirname);
    new karma.Server({
        files,
        frameworks: ['browserify', 'jasmine'],
        browsers: ['Chrome', 'Firefox', 'IE'],
        preprocessors: {
            "./test/**/*.js": ['browserify']
        },
        plugins: [
            'karma-browserify',
            'karma-chrome-launcher',
            'karma-firefox-launcher',
            'karma-ie-launcher',
            'karma-jasmine'
        ],
        browserify: {
            debug: true,
            transform: [["babelify", BABEL_CONFIG]],
            configure(bundle) {
                bundle.once('prebundle', () => {
                    bundle.transform('babelify').plugin('proxyquire-universal');
                });
            }
        },
        singleRun: true,
        autoWatch: false
    }, cb).start();
};

const runJasmine = (configPath, cb) => {
    process.chdir(__dirname);
    clearRequireCache();
    const jasmine = new Jasmine();
    jasmine.loadConfigFile(configPath);
    jasmine.onComplete(passed => cb(null));
    jasmine.execute();
};

gulp.task('browser', ['blank'], () => {
    return browserify({
        entries: PATHS.browserify.source,
        debug: true,
        standalone: BROWSERIFY_STANDALONE_NAME
    })
        .transform("babelify", BABEL_CONFIG)
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

gulp.task("unit", cb => {
    runJasmine(PATHS.jasmineConfigs.unit, cb);
});

gulp.task("e2e-generate", cb => {
    runJasmine(PATHS.jasmineConfigs.e2eGenerate, cb);
});

gulp.task("e2e-parse", cb => {
    runJasmine(PATHS.jasmineConfigs.e2eParse, cb);
});

gulp.task('e2e-browser', cb => {
    runKarma(["./test/helpers/**/*.js", "./browser/xlsx-populate.js", "./test/e2e-browser/**/*.spec.js"], cb);
});

gulp.task('unit-browser', cb => {
    runKarma(PATHS.karma, cb);
});

gulp.task("blank", () => {
    return Promise
        .all([
            fs.readFileAsync(PATHS.blank.workbook, "base64"),
            fs.readFileAsync(PATHS.blank.template, "utf8")
        ])
        .spread((data, template) => {
            const output = template.replace("{{DATA}}", data);
            return fs.writeFileAsync(PATHS.blank.build, output);
        });
});

gulp.task("docs", () => {
    return fs.readFileAsync(PATHS.readme.template, "utf8")
        .then(text => {
            const tocText = toc(text, { filter: str => str.indexOf('NOTOC-') === -1 }).content;
            text = text.replace("<!-- toc -->", tocText);
            text = text.replace(/NOTOC-/mg, "");
            return jsdoc2md.render({ files: PATHS.lib })
                .then(apiText => {
                    apiText = apiText.replace(/^#/mg, "##");
                    text = text.replace("<!-- api -->", apiText);
                    return fs.writeFileAsync(PATHS.readme.build, text);
                });
        });
});

gulp.task('watch', () => {
    // Only watch blank, unit, and docs for changes. Everything else is too slow or noisy.
    gulp.watch([PATHS.blank.template, PATHS.blank.workbook], ['blank']);
    gulp.watch(PATHS.unitTestSources, ["unit"]);
    gulp.watch([PATHS.lib, PATHS.readme.template], ["docs"]);
});

gulp.task('build', cb => {
    runSequence(["docs", "browser"], "lint", "unit", "unit-browser", "e2e-parse", "e2e-generate", "e2e-browser", cb);
});

gulp.task("default", cb => {
    // Watch just the quick stuff to aid development.
    runSequence("blank", ["unit", "docs"], "watch", cb);
});
