/**
 * Release version gate for the GitHub Actions pipeline (and local use).
 *
 * Verifies a push to main is release-ready:
 *   1. `config/package-solution.json` solution.version and `package.json`
 *      version agree (the .sppkg reports the former, npm the latter).
 *   2. If a latest release version is given (argv[2], tag without the leading
 *      "v"), the solution version must be GREATER than it — i.e. the version
 *      was actually bumped. Equal or lower => exit 1 and print why.
 *
 * Run:  node scripts/check-release-version.js [latestReleaseVersion]
 * Exit: 0 = release-ready, 1 = gate blocked
 */
'use strict';

const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');

function readVersion(relPath) {
  const raw = JSON.parse(fs.readFileSync(path.join(ROOT, relPath), 'utf8'));
  return relPath === 'package.json' ? raw.version : raw.solution.version;
}

/** true when a > b, dotted numeric comparison across any segment count. */
function greaterThan(a, b) {
  const as = String(a).trim().split('.').map((n) => parseInt(n, 10) || 0);
  const bs = String(b).trim().split('.').map((n) => parseInt(n, 10) || 0);
  const len = Math.max(as.length, bs.length);
  for (let i = 0; i < len; i++) {
    const x = as[i] || 0;
    const y = bs[i] || 0;
    if (x > y) return true;
    if (x < y) return false;
  }
  return false; // equal
}

const solutionVersion = readVersion('config/package-solution.json');
const packageVersion = readVersion('package.json');
const latest = (process.argv[2] || '').trim().replace(/^v/i, '');

const problems = [];
if (solutionVersion !== packageVersion) {
  problems.push(
    `Version drift: config/package-solution.json is ${solutionVersion} but package.json is ${packageVersion}. Keep them in sync before releasing.`
  );
}
if (latest && !greaterThan(solutionVersion, latest)) {
  problems.push(
    `Version not incremented: latest release is v${latest} but solution.version is ${solutionVersion}. Bump config/package-solution.json (and package.json to match) in this commit before pushing to main.`
  );
}

if (problems.length > 0) {
  console.error(
    `Release gate blocked (solution ${solutionVersion}, package.json ${packageVersion}, latest release v${latest || 'none'}):`
  );
  for (const p of problems) {
    console.error(`  - ${p}`);
  }
  process.exit(1);
}

console.log(
  `Release gate passed: solution ${solutionVersion}, package.json ${packageVersion}, latest release v${latest || 'none'}.`
);