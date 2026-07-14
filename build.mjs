// Build step for the Monthly Report Generator single-file app.
//
// Why this exists: the app used to ship ~500KB of JSX inside a
// <script type="text/babel"> tag and compile it in the browser with Babel
// Standalone on every page load. That made first paint slow (Babel had to
// parse + transpile the whole app before React could mount).
//
// This script moves the JSX source into app.jsx (the file you now edit) and
// pre-compiles it into plain JS that is inlined into index.html. The runtime
// no longer loads Babel. The output is still ONE self-contained index.html,
// served exactly as before (same origin, same localStorage key, same Firebase
// config) — so existing saved data is untouched.
//
// Usage:
//   node build.mjs      # or: npm run build
// Edit app.jsx, then run this to update index.html.
//
// The JSX transform is intentionally "classic" (React.createElement /
// React.Fragment, no automatic react/jsx-runtime import) to match the app's
// original Babel preset. Using the automatic runtime would inject an
// `import ... from "react/jsx-runtime"` into a classic script and black-screen
// the app (see PROJECT_RESUME.md §12).

import { readFileSync, writeFileSync, existsSync } from 'node:fs';
import { transformSync } from 'esbuild';

const HTML = new URL('./index.html', import.meta.url);
const SRC = new URL('./app.jsx', import.meta.url);

const START = '<!-- APP:BUNDLE:START (generated from app.jsx by build.mjs — do not edit between these markers) -->';
const END = '<!-- APP:BUNDLE:END -->';

let html = readFileSync(HTML, 'utf8');

// ---------------------------------------------------------------------------
// One-time migration: the first time this runs against the original file, pull
// the inline JSX out to app.jsx and strip the in-browser Babel machinery.
// ---------------------------------------------------------------------------
const BABEL_OPEN = '<script type="text/babel" data-presets="classic-react">';
if (html.includes(BABEL_OPEN)) {
  const openStart = html.indexOf(BABEL_OPEN);
  const bodyStart = openStart + BABEL_OPEN.length;
  const closeIdx = html.indexOf('</script>', bodyStart);
  if (closeIdx === -1) throw new Error('Could not find closing </script> for the app script.');

  const jsx = html.slice(bodyStart, closeIdx).replace(/^\r?\n/, '');

  if (!existsSync(SRC)) {
    writeFileSync(SRC, jsx, 'utf8');
    console.log(`• Extracted ${jsx.length.toLocaleString()} chars of JSX -> app.jsx`);
  } else {
    console.log('• app.jsx already exists — keeping it as the source of truth (ignoring inline copy).');
  }

  const before = html.slice(0, openStart);
  const after = html.slice(closeIdx + '</script>'.length);
  html = `${before}${START}\n  <script>\n  /* replaced by build */\n  </script>\n  ${END}${after}`;

  // Drop the Babel Standalone CDN <script> and the registerPreset <script>.
  html = html.replace(/[ \t]*<script src="https:\/\/unpkg\.com\/@babel\/standalone\/babel\.min\.js"><\/script>\r?\n/, '');
  html = html.replace(/[ \t]*<script>\s*Babel\.registerPreset\([\s\S]*?<\/script>\r?\n/, '');

  console.log('• Removed Babel Standalone CDN + registerPreset from index.html');
}

// ---------------------------------------------------------------------------
// Compile app.jsx -> plain JS and inline it between the markers.
// ---------------------------------------------------------------------------
if (!existsSync(SRC)) throw new Error('app.jsx not found and index.html has no inline JSX to extract.');

const source = readFileSync(SRC, 'utf8');
const { code, warnings } = transformSync(source, {
  loader: 'jsx',
  jsx: 'transform',
  jsxFactory: 'React.createElement',
  jsxFragment: 'React.Fragment',
  target: 'esnext', // don't down-level modern JS; the browser ran it as-is before
  charset: 'utf8',
});
for (const w of warnings) console.warn('  esbuild warning:', w.text);

// Defensive: keep any literal "</script" in the code from closing the tag.
const safe = code.replace(/<\/script/gi, '<\\/script');

const sIdx = html.indexOf(START);
const eIdx = html.indexOf(END);
if (sIdx === -1 || eIdx === -1) throw new Error('Bundle markers not found in index.html.');

const head = html.slice(0, sIdx + START.length);
const tail = html.slice(eIdx);
html = `${head}\n  <script>\n${safe}\n  </script>\n  ${tail}`;

writeFileSync(HTML, html, 'utf8');
console.log(`✓ Inlined ${safe.length.toLocaleString()} chars of compiled JS into index.html`);
