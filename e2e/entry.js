// End-to-end harness page: the real compiled app (lib/) talking HTTP to the SharePoint emulator (e2e/spEmulator.cjs).
// ?user=<email>&owner=1  who is signed in, and whether they own the site
// ?host=/sites/eng       the site the web part runs on (default /sites/hr)
// ?teams=1               running in Microsoft Teams (the site is found as the web part does: services/site.ts)
// ?site=<url>            the web part's "Site with the Time Off Manager lists" setting
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { initializeIcons } from '@fluentui/react';
import App from '../lib/webparts/timeOffManager/components/App';
import { LIST_PREFIX, listTitle } from '../lib/webparts/timeOffManager/schema/schema';
import { TEAMS_CHOICE_KEY, resolveSite } from '../lib/webparts/timeOffManager/services/site';
import { version as APP_VERSION } from '../package.json';

const params = new URLSearchParams(window.location.search);
const user = params.get('user') || 'admin@contoso.com';
const owner = params.get('owner') === '1';
const inTeams = params.get('teams') === '1';
const HOST = window.location.origin + (params.has('host') ? params.get('host') : '/sites/hr');

// SPHttpClient look-alike over fetch: same calls, same Response shape, plus the user header.
const call = (method) => (url, _config, options) => {
  const o = options || {};
  return fetch(url, { method, headers: { ...(o.headers || {}), 'X-User': user }, body: o.body });
};
const http = { get: call('GET'), post: call('POST') };
const getJson = async (url) => {
  const res = await http.get(url, undefined, { headers: { Accept: 'application/json;odata=nometadata' } });
  if (!res.ok) throw Object.assign(new Error(res.status + ''), { status: res.status });
  return res.json();
};
let remembered = '';
try {
  remembered = window.localStorage.getItem(TEAMS_CHOICE_KEY) || '';
} catch (e) {
  remembered = '';
}

initializeIcons(undefined, { disableWarnings: true });
resolveSite({ hostUrl: HOST, siteSetting: params.get('site') || '', inTeams, settingsTitle: listTitle(LIST_PREFIX, 'settings'), remembered, getJson }).then((site) => {
  window.__site = site;
  ReactDom.render(
    React.createElement(App, {
      http,
      webUrl: site.url,
      siteTitle: 'Human Resources',
      permissions: { manageLists: owner, manageWeb: owner },
      tenantId: '6f1c2a9e-0b7d-4c55-9e3a-2d8f41b07c13',
      prefix: LIST_PREFIX,
      appVersion: APP_VERSION,
      host: { inTeams, view: 'app', siteSource: site.source, siteError: site.error, siteChoices: site.choices }
    }),
    document.getElementById('root')
  );
});
