// Finding the Time Off Manager site (Teams) and the page used in email links: services/site.ts.
const test = require('node:test');
const assert = require('node:assert/strict');
const { lib } = require('./helpers.cjs');
const S = lib('services/site');

const O = 'https://contoso.sharepoint.com';

test('site addresses: pages and lists inside a site resolve to the site; other tenants are refused', () => {
  assert.equal(S.normalizeSiteUrl(`${O}/sites/hr/SitePages/Leave.aspx?x=1#/home`, O), `${O}/sites/hr`);
  assert.equal(S.normalizeSiteUrl('/sites/hr/Lists/BrillTM_Requests/AllItems.aspx', O), `${O}/sites/hr`);
  assert.equal(S.normalizeSiteUrl('https://fabrikam.sharepoint.com/sites/hr', O), '');
  assert.equal(S.normalizeSiteUrl('javascript:alert(1)', O), '');
});

test('email links: the real page without query string; never Teams’ host page or the workbench', () => {
  assert.equal(S.shareableBase(`${O}/sites/hr/SitePages/Leave.aspx?debug=true#/approvals`, ''), `${O}/sites/hr/SitePages/Leave.aspx`);
  assert.equal(S.shareableBase(`${O}/sites/eng/_layouts/15/teamshostedapp.aspx?teams`, `${O}/sites/hr/SitePages/Leave.aspx`), `${O}/sites/hr/SitePages/Leave.aspx`);
  assert.equal(S.isShareablePage(`${O}/_layouts/15/workbench.aspx`), false);
});

test('search results: only the settings list itself, each site once', () => {
  const row = (Title, Path, SPWebUrl) => ({ Cells: [{ Key: 'Title', Value: Title }, { Key: 'Path', Value: Path }, { Key: 'SPWebUrl', Value: SPWebUrl }] });
  const rows = [
    row('BrillTM_Settings', `${O}/sites/hr/Lists/BrillTM_Settings/AllItems.aspx`, `${O}/sites/hr`),
    row('Item', `${O}/sites/hr/Lists/BrillTM_Settings/DispForm.aspx?ID=1`, `${O}/sites/hr`),
    row('BrillTM_SettingsOld', `${O}/sites/old/Lists/BrillTM_SettingsOld`, `${O}/sites/old`),
    row('BrillTM_Settings', `https://fabrikam.sharepoint.com/sites/x/Lists/BrillTM_Settings`, 'https://fabrikam.sharepoint.com/sites/x'),
    row('BrillTM_Settings', `${O}/sites/emea/Lists/BrillTM_Settings`, `${O}/sites/emea`)
  ];
  assert.deepEqual(S.sitesFromSearch(rows, 'BrillTM_Settings', O), [`${O}/sites/hr`, `${O}/sites/emea`]);
});

test('several sites: this device’s choice, else the marked one, else the person chooses', () => {
  const f = [{ url: `${O}/sites/hr`, marked: false }, { url: `${O}/sites/emea`, marked: true }];
  assert.deepEqual(S.pickSite([f[0]], ''), { url: f[0].url });
  assert.deepEqual(S.pickSite(f, `${O}/sites/hr/`), { url: f[0].url });
  assert.deepEqual(S.pickSite(f, ''), { url: f[1].url });
  assert.deepEqual(S.pickSite(f.map((x) => ({ ...x, marked: false })), ''), { choose: [f[0].url, f[1].url] });
  assert.equal(S.pickSite([], ''), undefined);
});

test('resolveSite: page site, tab setting, Teams search, and outside Teams no search', async () => {
  const has = new Set([`${O}/sites/hr`]);
  const calls = [];
  const getJson = async (url) => {
    calls.push(url);
    const site = url.split('/_api/')[0];
    if (url.includes('/_api/search/')) return { PrimaryQueryResult: { RelevantResults: { Table: { Rows: [{ Cells: [{ Key: 'Title', Value: 'BrillTM_Settings' }, { Key: 'Path', Value: `${O}/sites/hr/Lists/BrillTM_Settings` }, { Key: 'SPWebUrl', Value: `${O}/sites/hr` }] }] } } } };
    if (url.includes("getbytitle('BrillTM_Settings')/items")) return { value: [] };
    if (url.includes("getbytitle('BrillTM_Settings')")) {
      if (has.has(site)) return { Id: 'x' };
      throw new Error('404');
    }
    throw new Error('unexpected ' + url);
  };
  const base = { settingsTitle: 'BrillTM_Settings', getJson };
  assert.deepEqual(await S.resolveSite({ ...base, hostUrl: `${O}/sites/hr`, inTeams: true }), { url: `${O}/sites/hr`, source: '' });
  assert.deepEqual(await S.resolveSite({ ...base, hostUrl: `${O}/sites/eng`, inTeams: true }), { url: `${O}/sites/hr`, source: 'found' });
  calls.length = 0;
  assert.deepEqual(await S.resolveSite({ ...base, hostUrl: `${O}/sites/eng`, inTeams: false }), { url: `${O}/sites/eng`, source: '' });
  assert.ok(!calls.some((c) => c.includes('/search/')), 'no search outside Teams');
  assert.deepEqual(await S.resolveSite({ ...base, hostUrl: `${O}/sites/eng`, inTeams: false, siteSetting: '/sites/hr/SitePages/Home.aspx' }), { url: `${O}/sites/hr`, source: 'property' });
  assert.match((await S.resolveSite({ ...base, hostUrl: `${O}/sites/eng`, inTeams: true, siteSetting: 'https://evil.example.com' })).error, /not a SharePoint site address/);
});
