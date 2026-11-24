// Every Fluent UI icon name the app uses exists in the icon font (an unknown name renders as a blank square).
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');

function knownIcons() {
  const dir = path.join(root, 'node_modules/@fluentui/font-icons-mdl2/lib-commonjs');
  const known = new Set();
  for (const f of fs.readdirSync(dir)) {
    if (!/^fabric-icons.*\.js$/.test(f)) continue;
    for (const m of fs.readFileSync(path.join(dir, f), 'utf8').matchAll(/'?([A-Za-z0-9]+)'?:\s*'\\u[0-9A-Fa-f]{4}'/g)) known.add(m[1]);
  }
  for (const m of fs.readFileSync(path.join(dir, 'iconAliases.js'), 'utf8').matchAll(/'?([A-Za-z0-9]+)'?:\s*'/g)) known.add(m[1]);
  return known;
}

function usedIcons() {
  const used = new Map();
  const walk = (d) => {
    for (const f of fs.readdirSync(d)) {
      const p = path.join(d, f);
      if (fs.statSync(p).isDirectory()) walk(p);
      else if (/\.tsx?$/.test(f)) {
        const s = fs.readFileSync(p, 'utf8');
        const patterns = [/iconName[=:]\s*\{?\s*['"]([A-Za-z0-9]+)['"]/g, /\bicon[=:]\s*\{?\s*['"]([A-Za-z0-9]+)['"]/g, /iconName=\{[^}]*\}/g];
        for (const re of patterns.slice(0, 2)) for (const m of s.matchAll(re)) used.set(m[1], path.relative(root, p));
        // Names chosen in an expression: iconName={x ? 'A' : 'B'}
        for (const m of s.matchAll(patterns[2])) for (const q of m[0].matchAll(/'([A-Z][A-Za-z0-9]+)'/g)) used.set(q[1], path.relative(root, p));
      }
    }
  };
  walk(path.join(root, 'src'));
  return used;
}

test('icons: every icon name in the code (and the 30 leave types) exists in the Fluent UI icon font', () => {
  const known = knownIcons();
  assert.ok(known.size > 1000, 'icon registry read');
  const used = usedIcons();
  assert.ok(used.size > 50, `${used.size} icon names found`);
  const missing = [...used].filter(([n]) => !known.has(n)).map(([n, f]) => `${n} (${f})`);
  assert.deepEqual(missing, []);
});
