/**
 * Build-time generator: reads docs/configuration.md and emits a compact TS
 * module (help/helpContent.g.ts) containing the markdown as a string constant,
 * so the in-app "?" help dialog renders the same reference shipped in the
 * repository — a single source of truth, like generateDemoData.js.
 *
 * Run:  node scripts/generateHelpContent.js  (Node 18)
 * Output: src/webparts/sharepointPowerCharts/help/helpContent.g.ts
 */
'use strict';

const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const SRC_MD = path.join(ROOT, 'docs', 'configuration.md');
const OUT_TS = path.join(
  ROOT,
  'src',
  'webparts',
  'sharepointPowerCharts',
  'help',
  'helpContent.g.ts'
);

const markdown = fs.readFileSync(SRC_MD, 'utf8');

const header = [
  '// GENERATED FILE - do not edit by hand.',
  '// Regenerate with: node scripts/generateHelpContent.js',
  '// Source of truth: docs/configuration.md',
  ''
].join('\n');

const content = header + 'export const HELP_MARKDOWN: string = ' + JSON.stringify(markdown) + ';\n';
fs.writeFileSync(OUT_TS, content, 'utf8');
console.log(`Wrote ${OUT_TS} (${markdown.length} chars from ${SRC_MD})`);