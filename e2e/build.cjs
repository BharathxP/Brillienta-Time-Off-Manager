// Builds the e2e page (e2e-dist/) from the compiled app in lib/ (run `npm run build:dev` first).
const fs = require('fs');
const path = require('path');
const webpack = require('webpack');
const base = require('../preview/webpack.config.cjs');

const out = path.resolve(__dirname, '../e2e-dist');
fs.mkdirSync(out, { recursive: true });
fs.writeFileSync(path.join(out, 'index.html'), '<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>Time Off Manager e2e</title><style>body{margin:0;padding:12px;background:#f3f5f4}</style></head><body><div id="root"></div><script src="/bundle.js"></script></body></html>');
webpack({ ...base, entry: path.resolve(__dirname, 'entry.js'), output: { path: out, filename: 'bundle.js' } }, (err, stats) => {
  if (err || stats.hasErrors()) {
    console.error(err || stats.toString({ all: false, errors: true }));
    process.exit(1);
  }
  console.log('e2e page built: ' + out);
});
