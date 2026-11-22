// Tests run against the compiled CommonJS output (npm run build:dev first).
const path = require('path');
const lib = (p) => require(path.join(__dirname, '..', 'lib-commonjs', 'webparts', 'timeOffManager', p));
module.exports = { lib };
