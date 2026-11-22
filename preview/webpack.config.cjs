// Local preview bundle (npm run preview): the compiled app from lib/ with SharePoint replaced by an in-memory fake.
const path = require('path');

module.exports = {
  mode: 'development',
  devtool: false,
  entry: path.resolve(__dirname, 'entry.js'),
  output: { path: path.resolve(__dirname, '../preview-dist'), filename: 'bundle.js' },
  resolve: {
    alias: {
      // SPFx's HTTP client only matters for its configuration constant here.
      '@microsoft/sp-http$': path.resolve(__dirname, 'spHttpStub.js'),
      './TimeOff.module.scss$': path.resolve(__dirname, 'styles.js')
    }
  },
  module: {
    rules: [{ test: /\.css$/, type: 'asset/source' }]
  },
  performance: { hints: false }
};
