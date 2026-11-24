// Publication: what Microsoft Marketplace (AppSource) validation and Teams "Sync to Teams" need from the package.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const manifest = JSON.parse(fs.readFileSync(path.join(root, 'src/webparts/timeOffManager/TimeOffManagerWebPart.manifest.json'), 'utf8'));
const pkg = JSON.parse(fs.readFileSync(path.join(root, 'config/package-solution.json'), 'utf8')).solution;
const version = require('../package.json').version;

function pngSize(file) {
  const b = fs.readFileSync(file);
  assert.equal(b.toString('ascii', 1, 4), 'PNG', `${file} is a PNG`);
  return { w: b.readUInt32BE(16), h: b.readUInt32BE(20), colorType: b[25] };
}

/** File names in a zip (central directory). */
function zipNames(buf) {
  const names = [];
  for (let i = 0; i < buf.length - 4; i++) {
    if (buf.readUInt32LE(i) !== 0x02014b50) continue;
    const len = buf.readUInt16LE(i + 28);
    names.push(buf.toString('utf8', i + 46, i + 46 + len));
  }
  return names;
}

test('package: version is the same everywhere', () => {
  assert.equal(pkg.version, `${version}.0`);
  assert.equal(pkg.features[0].version, `${version}.0`);
});

test('package: Teams hosts declared and tenant-wide deployment (required by Sync to Teams)', () => {
  assert.ok(manifest.supportedHosts.includes('TeamsTab'));
  assert.ok(manifest.supportedHosts.includes('TeamsPersonalApp'));
  assert.equal(pkg.skipFeatureDeployment, true);
});

test('package: Teams and marketplace text limits', () => {
  for (const e of manifest.preconfiguredEntries) {
    assert.ok(e.title.default.length <= 30, `${e.title.default}: title ≤ 30`);
    assert.ok(e.description.default.length <= 80, `${e.title.default}: description ≤ 80`);
  }
  assert.ok(pkg.metadata.shortDescription.default.length <= 80, 'short description ≤ 80');
  assert.ok(pkg.metadata.longDescription.default.length <= 4000, 'long description ≤ 4000');
});

test('marketplace: picker icon, app name and description links match the listing', () => {
  const icon = 'data:image/png;base64,' + fs.readFileSync(path.join(root, 'sharepoint', pkg.iconPath)).toString('base64');
  for (const e of manifest.preconfiguredEntries) {
    assert.equal(e.iconImageUrl, icon, `${e.title.default} uses the app icon in the picker`);
    assert.equal(e.officeFabricIconFontName, undefined, `${e.title.default}: no font icon`);
  }
  assert.equal(pkg.name, 'Brillienta Time Off Manager', 'same name as the marketplace listing');
  const d = pkg.metadata.longDescription.default;
  for (const s of ['Get started', 'Help', 'Contact us', 'Sync to Teams', 'https://brillienta.com/time-off-manager', 'https://kb.brillienta.com/time-off-manager/', 'https://brillienta.com/contact', 'support@brillienta.com']) assert.ok(d.includes(s), s);
});

test('marketplace: app icon 96 px and five 1366×768 screenshots', () => {
  assert.deepEqual(Object.values(pngSize(path.join(root, 'sharepoint', pkg.iconPath))).slice(0, 2), [96, 96]);
  assert.equal(pkg.metadata.screenshotPaths.length, 5);
  for (const s of pkg.metadata.screenshotPaths) {
    const { w, h } = pngSize(path.join(root, 'sharepoint', s));
    assert.deepEqual([w, h], [1366, 768], s);
  }
});

test('package: developer details (Teams rejects an app without them)', () => {
  const d = pkg.developer;
  assert.ok(d.name && d.name.length <= 32);
  for (const k of ['websiteUrl', 'privacyUrl', 'termsOfUseUrl']) assert.match(d[k], /^https:\/\/\S+$/, k);
  assert.ok(d.mpnId);
});

test('package: Graph permissions are the ones the app uses, and the description names them', () => {
  const scopes = pkg.webApiPermissionRequests.map((p) => p.scope).sort();
  assert.deepEqual(scopes, ['Calendars.ReadWrite', 'Group.Read.All', 'Group.ReadWrite.All', 'MailboxSettings.ReadWrite', 'User.Read.All']);
  for (const s of scopes) assert.ok(pkg.metadata.longDescription.default.includes(s), `description mentions ${s}`);
});

test('Teams app: teams/manifest.json is current and passes the official schema and app package rules', () => {
  const T = require('../tools/teams/app-manifest.cjs');
  const text = fs.readFileSync(path.join(root, 'teams', 'manifest.json'), 'utf8');
  assert.equal(text, T.manifestText(), 'regenerate with: npm run teams');
  const shipped = JSON.parse(text);
  const icons = T.iconNames(manifest.id);
  assert.deepEqual(shipped.icons, icons);
  const color = fs.readFileSync(path.join(root, 'teams', icons.color));
  const outline = fs.readFileSync(path.join(root, 'teams', icons.outline));
  assert.deepEqual(T.checkPackage(shipped, color, outline), []);
  assert.equal(shipped.name.short, pkg.name);
  assert.equal(shipped.version, version);
  assert.equal(shipped.id, manifest.id);
  assert.equal(shipped.supportsChannelFeatures, 'tier1');
  for (const t of shipped.staticTabs.concat(shipped.configurableTabs)) assert.ok((t.contentUrl || t.configurationUrl).includes(`componentId=${manifest.id}`));
  assert.deepEqual(zipNames(T.buildPackage()).sort(), [icons.color, icons.outline, 'manifest.json'].sort());
  // Only the Teams files in teams/ (the packager copies the whole folder into the .sppkg).
  assert.deepEqual(fs.readdirSync(path.join(root, 'teams')).sort(), [icons.color, icons.outline, 'manifest.json'].sort());
});

test('publication documents exist: privacy statement and terms', () => {
  for (const f of ['docs/PRIVACY.md', 'docs/TERMS.md']) assert.ok(fs.readFileSync(path.join(root, f), 'utf8').length > 300, f);
});
