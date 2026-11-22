#!/usr/bin/env node
// Brillienta Time Off Manager licence tool (vendor only; never shipped in the .sppkg).
//
// Uses the same Brillienta signing key pair as the Employee Directory (tools/license in the repository root):
// the public half is src/webparts/timeOffManager/services/licensing/publicKey.ts.
//
//   node tools/license/license.cjs issue --tenant <tenant id> --customer "Contoso Ltd" --users 120
//        [--expires 2027-12-31] [--support unlimited] [--key <private-key.pem>] [--id <licence id>]
//   node tools/license/license.cjs verify <licence key> [--tenant <tenant id>]
//
// Keys: ECDSA P-256 / SHA-256. Format: TO1.<base64url(payload JSON)>.<base64url(signature, IEEE P1363)>

const crypto = require('crypto');
const fs = require('fs');
const os = require('os');
const path = require('path');

const PUBLIC_KEY_FILE = path.join(__dirname, '..', '..', 'src', 'webparts', 'timeOffManager', 'services', 'licensing', 'publicKey.ts');
const DEFAULT_KEY = path.join(os.homedir(), '.employee-directory-licensing', 'private-key.pem');
const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const b64url = (buf) => Buffer.from(buf).toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
const fromB64url = (s) => Buffer.from(s.replace(/-/g, '+').replace(/_/g, '/'), 'base64');

function args(argv) {
  const out = { _: [] };
  for (let i = 0; i < argv.length; i++) {
    if (argv[i].startsWith('--')) out[argv[i].slice(2)] = argv[i + 1] && !argv[i + 1].startsWith('--') ? argv[++i] : 'true';
    else out._.push(argv[i]);
  }
  return out;
}
function fail(msg) {
  console.error('Error: ' + msg);
  process.exit(1);
}
function appPublicKey() {
  const m = /LICENSE_PUBLIC_KEY\s*=\s*'([^']+)'/.exec(fs.readFileSync(PUBLIC_KEY_FILE, 'utf8'));
  if (!m) fail('No public key in ' + PUBLIC_KEY_FILE);
  return crypto.createPublicKey({ key: Buffer.from(m[1], 'base64'), format: 'der', type: 'spki' });
}

const a = args(process.argv.slice(2));
const cmd = a._[0];
if (cmd === 'issue') {
  if (!a.tenant || !GUID.test(a.tenant)) fail('--tenant <tenant id GUID> is required (Admin › Licence shows it).');
  if (!a.customer) fail('--customer is required.');
  const users = Number(a.users);
  if (!(users > 0)) fail('--users <number> is required.');
  if (a.expires && !/^\d{4}-\d{2}-\d{2}$/.test(a.expires)) fail('--expires must be yyyy-mm-dd.');
  const keyFile = a.key || DEFAULT_KEY;
  if (!fs.existsSync(keyFile)) fail('Private key not found: ' + keyFile);
  const privateKey = crypto.createPrivateKey(fs.readFileSync(keyFile));
  const payload = { v: 1, id: a.id || 'TO-' + crypto.randomBytes(4).toString('hex').toUpperCase(), product: 'timeoff', tenant: a.tenant.toLowerCase(), customer: a.customer, users, issued: new Date().toISOString().substring(0, 10) };
  if (a.expires) payload.expires = a.expires;
  if (a.support) payload.support = a.support;
  const signed = 'TO1.' + b64url(JSON.stringify(payload));
  const sig = crypto.sign('sha256', Buffer.from(signed), { key: privateKey, dsaEncoding: 'ieee-p1363' });
  const key = signed + '.' + b64url(sig);
  if (!crypto.verify('sha256', Buffer.from(signed), { key: appPublicKey(), dsaEncoding: 'ieee-p1363' }, sig)) fail('This private key does not match the public key in the app.');
  console.log(key);
} else if (cmd === 'verify') {
  const parts = (a._[1] || '').replace(/\s+/g, '').split('.');
  if (parts.length !== 3 || parts[0] !== 'TO1') fail('Not a Time Off Manager key.');
  const ok = crypto.verify('sha256', Buffer.from(parts[0] + '.' + parts[1]), { key: appPublicKey(), dsaEncoding: 'ieee-p1363' }, fromB64url(parts[2]));
  const payload = JSON.parse(fromB64url(parts[1]).toString('utf8'));
  console.log(JSON.stringify(payload, null, 2));
  console.log(ok ? 'Signature: valid' : 'Signature: INVALID');
  if (a.tenant && a.tenant.toLowerCase() !== payload.tenant) console.log('Tenant: does NOT match ' + a.tenant);
  process.exit(ok ? 0 : 1);
} else {
  console.log(fs.readFileSync(__filename, 'utf8').split('\n').slice(1, 12).join('\n').replace(/^\/\/ ?/gm, ''));
}
