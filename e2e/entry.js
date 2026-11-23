// End-to-end harness page: the real compiled app (lib/) talking HTTP to the SharePoint emulator (e2e/spEmulator.cjs).
// ?user=<email>&owner=1 — who is signed in, and whether they own the site (Manage Lists / Manage Web).
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { initializeIcons } from '@fluentui/react';
import App from '../lib/webparts/timeOffManager/components/App';
import { version as APP_VERSION } from '../package.json';

const params = new URLSearchParams(window.location.search);
const user = params.get('user') || 'admin@contoso.com';
const owner = params.get('owner') === '1';
const WEB = window.location.origin + '/sites/hr';

// SPHttpClient look-alike over fetch: same calls, same Response shape, plus the user header.
const call = (method) => (url, _config, options) => {
  const o = options || {};
  return fetch(url, { method, headers: { ...(o.headers || {}), 'X-User': user }, body: o.body });
};
const http = { get: call('GET'), post: call('POST') };

initializeIcons(undefined, { disableWarnings: true });
ReactDom.render(
  React.createElement(App, {
    http,
    webUrl: WEB,
    siteTitle: 'Human Resources',
    permissions: { manageLists: owner, manageWeb: owner },
    tenantId: '6f1c2a9e-0b7d-4c55-9e3a-2d8f41b07c13',
    prefix: 'BrillTM_',
    appVersion: APP_VERSION,
    host: { inTeams: false, view: 'app' }
  }),
  document.getElementById('root')
);
