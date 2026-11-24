// The Microsoft Teams app manifest shipped in the package as teams/manifest.json, next to the icons
// teams/<web part id>_color.png and _outline.png: the same layout as Brillienta Employee Directory and Helpdesk.
// The packager copies everything in teams/ into the .sppkg, and Sync to Teams deploys this manifest instead of
// generating one from the web part (a generated one is named after the web part title and lacks
// supportsChannelFeatures, which Teams now requires). Everything is taken from config/package-solution.json, the web
// part manifest, package.json and the icons, so the Teams app always matches the SharePoint package.
//
// buildManifest()  → the manifest object
// buildPackage()   → a zip with manifest.json and the icons at its root, for uploading in Teams admin center
// checkPackage()   → a list of problems against the Teams manifest rules (empty when valid)
//
// Written by `node tools/teams/make-app-package.cjs`; checked by tests/teams.test.cjs.
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..', '..');
const read = (p) => fs.readFileSync(path.join(root, p));
const json = (p) => JSON.parse(read(p).toString('utf8'));

// 1.25 or later: Teams no longer accepts apps with channel tabs that don't declare supportsChannelFeatures, a
// v1.25 property ("Update your manifest to include supportsChannelFeatures=tier1").
const MANIFEST_VERSION = '1.25';
const SCHEMA = `https://developer.microsoft.com/en-us/json-schemas/teams/v${MANIFEST_VERSION}/MicrosoftTeams.schema.json`;
/** SharePoint Online's first-party app id: required for SharePoint Framework apps (Microsoft's SPFx Teams docs). */
const SHAREPOINT_APP_ID = '00000003-0000-0ff1-ce00-000000000000';
const ACCENT = '#038F62';

/** Icon file names in teams/ (and in the manifest): named after the web part, as SharePoint expects. */
function iconNames(id) {
  return { color: `${id}_color.png`, outline: `${id}_outline.png` };
}

function sources() {
  const webPart = json('src/webparts/timeOffManager/TimeOffManagerWebPart.manifest.json');
  return {
    webPart,
    pkg: json('config/package-solution.json').solution,
    version: json('package.json').version,
    color: read(`teams/${iconNames(webPart.id).color}`),
    outline: read(`teams/${iconNames(webPart.id).outline}`)
  };
}

function buildManifest(s = sources()) {
  const { webPart, pkg, version } = s;
  const id = webPart.id;
  // URLs from "Developer-provided Microsoft Teams app manifest & package" (Microsoft's SPFx docs); Teams fills in
  // {teamSiteDomain}, {teamSitePath} and {locale}.
  const personal = `https://{teamSiteDomain}/_layouts/15/TeamsLogon.aspx?SPFX=true&dest=/_layouts/15/teamshostedapp.aspx%3Fteams%26personal%26componentId=${id}%26forceLocale={locale}`;
  const tab = `https://{teamSiteDomain}{teamSitePath}/_layouts/15/TeamsLogon.aspx?SPFX=true&dest={teamSitePath}/_layouts/15/teamshostedapp.aspx%3FopenPropertyPane=true%26teams%26componentId=${id}%26forceLocale={locale}`;
  return {
    $schema: SCHEMA,
    manifestVersion: MANIFEST_VERSION,
    version,
    // The web part id, as Sync to Teams used before this package existed, so tenants that already synced get an
    // update rather than a second app.
    id,
    developer: {
      name: pkg.developer.name,
      websiteUrl: pkg.developer.websiteUrl,
      privacyUrl: pkg.developer.privacyUrl,
      termsOfUseUrl: pkg.developer.termsOfUseUrl,
      mpnId: pkg.developer.mpnId
    },
    name: { short: pkg.name, full: pkg.name },
    description: { short: pkg.metadata.shortDescription.default, full: pkg.metadata.longDescription.default },
    icons: iconNames(id),
    accentColor: ACCENT,
    staticTabs: [
      { entityId: id, name: webPart.preconfiguredEntries[0].title.default, contentUrl: personal, websiteUrl: pkg.developer.websiteUrl, scopes: ['personal'] }
    ],
    configurableTabs: [{ configurationUrl: tab, canUpdateConfiguration: true, scopes: ['team'], context: ['channelTab'] }],
    // Works in standard, private and shared channels: the tab shows Time Off Manager and doesn't use channel or team
    // membership, so tier1 needs nothing more (Microsoft's "Build apps for shared and private channels").
    supportsChannelFeatures: 'tier1',
    permissions: ['identity'],
    validDomains: [
      '*.login.microsoftonline.com',
      '*.sharepoint.com',
      '*.sharepoint-df.com',
      'spoppe-a.akamaihd.net',
      'spoprod-a.akamaihd.net',
      'resourceseng.blob.core.windows.net',
      'msft.spoppe.com'
    ],
    webApplicationInfo: { resource: 'https://{teamSiteDomain}', id: SHAREPOINT_APP_ID }
  };
}

/** Fixed date, so the zip only changes when its content does. */
const ZIP_DATE = new Date(2026, 0, 1, 0, 0, 0);

function buildPackage(s = sources()) {
  const { zipFiles } = require(path.join(root, 'lib-commonjs/webparts/timeOffManager/services/zip.js'));
  return Buffer.from(
    zipFiles(
      [
        { path: 'manifest.json', content: manifestText(s) },
        { path: iconNames(s.webPart.id).color, content: new Uint8Array(s.color) },
        { path: iconNames(s.webPart.id).outline, content: new Uint8Array(s.outline) }
      ],
      ZIP_DATE
    )
  );
}

/** Reads a PNG's size, colour type and decoded RGBA pixels (8-bit, non-interlaced: what our icons are). */
function png(buf) {
  const zlib = require('zlib');
  if (buf.toString('ascii', 1, 4) !== 'PNG') throw new Error('not a PNG');
  let at = 8;
  let w = 0;
  let h = 0;
  let type = 0;
  const idat = [];
  while (at < buf.length) {
    const len = buf.readUInt32BE(at);
    const kind = buf.toString('ascii', at + 4, at + 8);
    const data = buf.subarray(at + 8, at + 8 + len);
    if (kind === 'IHDR') {
      w = data.readUInt32BE(0);
      h = data.readUInt32BE(4);
      type = data[9];
      if (data[8] !== 8 || data[12] !== 0) throw new Error('only 8-bit non-interlaced PNGs are read');
    } else if (kind === 'IDAT') idat.push(data);
    at += 12 + len;
  }
  const ch = { 2: 3, 6: 4 }[type];
  if (!ch) throw new Error(`PNG colour type ${type} not read`);
  const raw = zlib.inflateSync(Buffer.concat(idat));
  const stride = w * ch;
  const px = Buffer.alloc(w * h * 4);
  let prev = Buffer.alloc(stride);
  for (let y = 0; y < h; y++) {
    const f = raw[y * (stride + 1)];
    const line = Buffer.from(raw.subarray(y * (stride + 1) + 1, (y + 1) * (stride + 1)));
    for (let i = 0; i < stride; i++) {
      const a = i >= ch ? line[i - ch] : 0;
      const b = prev[i];
      const c = i >= ch ? prev[i - ch] : 0;
      const p = a + b - c;
      const pa = Math.abs(p - a);
      const pb = Math.abs(p - b);
      const pc = Math.abs(p - c);
      const paeth = pa <= pb && pa <= pc ? a : pb <= pc ? b : c;
      line[i] = (line[i] + [0, a, b, (a + b) >> 1, paeth][f]) & 0xff;
    }
    for (let x = 0; x < w; x++) {
      const o = (y * w + x) * 4;
      px[o] = line[x * ch];
      px[o + 1] = line[x * ch + 1];
      px[o + 2] = line[x * ch + 2];
      px[o + 3] = ch === 4 ? line[x * ch + 3] : 255;
    }
    prev = line;
  }
  return { w, h, type, px };
}

/**
 * Problems with the manifest and icons against the Teams rules: the official JSON schema for the manifest version
 * (from @microsoft/app-manifest) plus the app package rules the schema doesn't express. Empty when valid.
 */
function checkPackage(manifest, color, outline) {
  const problems = [];
  const Ajv = require('ajv-draft-04');
  const addFormats = require('ajv-formats');
  const schema = require(`@microsoft/app-manifest/build/json-schemas/teams/v${MANIFEST_VERSION}/MicrosoftTeams.schema.json`);
  const ajv = new Ajv({ allErrors: true, strict: false });
  addFormats(ajv);
  const validate = ajv.compile(schema);
  if (!validate(manifest)) for (const e of validate.errors) problems.push(`schema: ${e.instancePath || '/'} ${e.message}${e.params && e.params.additionalProperty ? ` (${e.params.additionalProperty})` : ''}`);

  if (!/^\d+\.\d+\.\d+$/.test(manifest.version)) problems.push('version must be major.minor.patch');
  if (manifest.name.short.length > 30) problems.push('name.short over 30 characters');
  if (manifest.description.full === manifest.description.short) problems.push('description.full repeats description.short');
  for (const d of manifest.validDomains) if (d === '*' || /^\*\.[^.]+$/.test(d)) problems.push(`validDomains: ${d} is too broad`);
  if ((manifest.configurableTabs || []).length && manifest.supportsChannelFeatures !== 'tier1') problems.push('channel tabs need supportsChannelFeatures: tier1 (Teams rejects the upload without it)');
  if (manifest.supportedChannelTypes) problems.push('supportedChannelTypes is not accepted in manifest v1.25+, use supportsChannelFeatures');
  if (!manifest.webApplicationInfo || manifest.webApplicationInfo.id !== SHAREPOINT_APP_ID) problems.push('webApplicationInfo.id must be SharePoint Online');

  const c = png(color);
  if (c.w !== 192 || c.h !== 192) problems.push(`colour icon is ${c.w}×${c.h}, must be 192×192`);
  const o = png(outline);
  if (o.w !== 32 || o.h !== 32) problems.push(`outline icon is ${o.w}×${o.h}, must be 32×32`);
  if (o.type !== 6) problems.push('outline icon needs an alpha channel');
  let opaque = 0;
  for (let i = 0; i < o.px.length; i += 4) {
    const a = o.px[i + 3];
    if (a === 0) continue;
    opaque++;
    if (o.px[i] < 250 || o.px[i + 1] < 250 || o.px[i + 2] < 250) {
      problems.push('outline icon must be white on transparent');
      break;
    }
  }
  if (!opaque) problems.push('outline icon is empty');
  return problems;
}

/** The manifest as written to teams/manifest.json. */
function manifestText(s = sources()) {
  return JSON.stringify(buildManifest(s), null, 2) + '\n';
}

module.exports = { buildManifest, buildPackage, checkPackage, manifestText, iconNames, sources, png, MANIFEST_VERSION };
