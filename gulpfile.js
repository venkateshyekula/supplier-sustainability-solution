/* eslint-disable */
'use strict';

const build = require('@microsoft/sp-build-web');
const fs = require('fs');
const path = require('path');

// Suppress ms-Grid warning
build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

// Suppress filename module warning for SCSS
build.addSuppression(/filename should end with module/);

// Override serve task with serve-deprecated
var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);
  result.set('serve', result.get('serve-deprecated'));
  return result;
};

// Copy image assets from src to lib before bundling
const copyAssetsTask = build.subTask('copy-assets-task', (gulp, buildOptions, done) => {
  return gulp.src('src/**/*.{png,jpg,jpeg,gif,svg}')
    .pipe(gulp.dest('lib'));
});
build.rig.addPreBuildTask(copyAssetsTask);

// Synchronize package.json version with package-solution.json
const syncVersionTask = build.subTask('sync-version-task', (gulp, buildOptions, done) => {
  const packageJsonPath = path.join(__dirname, 'package.json');
  const packageSolutionPath = path.join(__dirname, 'config', 'package-solution.json');

  if (fs.existsSync(packageJsonPath) && fs.existsSync(packageSolutionPath)) {
    const packageJson = JSON.parse(fs.readFileSync(packageJsonPath, 'utf8'));
    const packageSolution = JSON.parse(fs.readFileSync(packageSolutionPath, 'utf8'));

    const npmVersion = packageJson.version || '1.0.0';
    const cleanVersion = npmVersion.split('-')[0];
    const parts = cleanVersion.split('.');

    const major = parseInt(parts[0], 10) || 0;
    const minor = parseInt(parts[1], 10) || 0;
    const patch = parseInt(parts[2], 10) || 0;
    const spfxVersion = `${major}.${minor}.${patch}.0`;

    packageSolution.solution.version = spfxVersion;

    if (Array.isArray(packageSolution.solution.features)) {
      packageSolution.solution.features.forEach((feature) => {
        feature.version = spfxVersion;
      });
    }

    fs.writeFileSync(packageSolutionPath, JSON.stringify(packageSolution, null, 2) + '\n', 'utf8');
    console.log(`[sync-version] package.json: ${npmVersion} -> package-solution.json: ${spfxVersion}`);
  }

  return Promise.resolve();
});
build.rig.addPreBuildTask(syncVersionTask);

build.initialize(require('gulp'));