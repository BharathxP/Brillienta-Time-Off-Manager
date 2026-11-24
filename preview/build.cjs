// Builds the local preview (npm run preview). Uses webpack's Node API, so no extra CLI package is needed.
const fs = require('fs');
const path = require('path');
const webpack = require('webpack');
const config = require('./webpack.config.cjs');

const out = path.resolve(__dirname, '../preview-dist');
fs.mkdirSync(out, { recursive: true });
['index.html'].forEach((f) => fs.copyFileSync(path.resolve(__dirname, f), path.join(out, f)));
const fonts = path.resolve(__dirname, '../node_modules/@fluentui/font-icons-mdl2/fonts');
fs.mkdirSync(path.join(out, 'fonts'), { recursive: true });
fs.readdirSync(fonts).forEach((f) => fs.copyFileSync(path.join(fonts, f), path.join(out, 'fonts', f)));

webpack(config, (err, stats) => {
  if (err || stats.hasErrors()) {
    console.error(err || stats.toString({ colors: false, all: false, errors: true }));
    process.exit(1);
  }
  console.log(`Preview built: ${path.join(out, 'index.html')}`);
  console.log('Open it in a browser (file://). Add ?role=employee, ?role=manager or ?role=hr to see other roles.');
});
