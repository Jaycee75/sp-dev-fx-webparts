'use strict';

// check if gulp dist was called
if (process.argv.indexOf('dist') !== -1) {
  // add ship options to command call
  process.argv.push('--ship');
}

const gulp = require('gulp');
const build = require('@microsoft/sp-build-web');
const gulpSequence = require('gulp-sequence');
build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

// Create clean distrubution package
gulp.task('dist', gulpSequence('clean', 'bundle', 'package-solution'));
// Create clean development package
gulp.task('dev', gulpSequence('clean', 'bundle', 'package-solution'));

/********************************************************************************************
 * Adds an alias for handlebars in order to avoid errors while gulping the project
 * https://github.com/wycats/handlebars.js/issues/1174
 * Adds a loader and a node setting for webpacking the handlebars-helpers correctly
 * https://github.com/helpers/handlebars-helpers/issues/263
 ********************************************************************************************/
build.configureWebpack.mergeConfig({
  additionalConfiguration: (generatedConfiguration) => {

    generatedConfiguration.resolve.alias = {
      ...(generatedConfiguration.resolve.alias || {}),
      handlebars: 'handlebars/dist/handlebars.min.js'
    };

    // handlebars-helpers (and create-frame) use lazy-cache style `require(name, alias)` calls
    // that webpack cannot statically analyse; unlazy-loader rewrites them into plain requires.
    generatedConfiguration.module.rules.push({
      test: /node_modules[\\/](handlebars-helpers|create-frame)[\\/].*\.js$/,
      loader: 'unlazy-loader'
    });

    // webpack 5 (used by @microsoft/sp-build-web 1.22.x) no longer polyfills Node core modules
    // and rejects the old `node: { fs: 'empty' }` setting, so configure resolve.fallback instead.
    generatedConfiguration.resolve.fallback = {
      ...(generatedConfiguration.resolve.fallback || {}),
      fs: false,
      path: require.resolve('path-browserify'),
      url: require.resolve('url/')
    };

    return generatedConfiguration;
  }
});


var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};

build.initialize(gulp);
