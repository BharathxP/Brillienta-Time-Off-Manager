// A strict stand-in for the SharePoint Online REST API, for end-to-end tests (npm run e2e).
//
// Unlike the preview's in-memory fake, it behaves like SharePoint where it matters for this app:
//  * a list must be created (POST web/lists) before it can be read or written — otherwise 404;
//  * a column must exist (fields/createfieldasxml) before an item can carry it — otherwise 400, as SharePoint does;
//  * Text columns hold at most 255 characters; Number / Boolean / DateTime values are type checked;
//  * $filter on a column that does not exist is a 400;
//  * every request carries a user (X-User header): currentuser, Author and Editor follow it, so several people can use
//    the same site at once from separate browser sessions;
//  * responses carry SharePoint's Date header (the app reads the server clock).
// State lives in memory for the life of the server; GET /__state returns it, POST /__reset clears it.

const http = require('http');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const SITE = '/sites/hr';
const BUILT_IN = { Id: 'Counter', ID: 'Counter', Title: 'Text', Created: 'DateTime', Modified: 'DateTime', Attachments: 'Boolean', AuthorId: 'Number', EditorId: 'Number' };

function newSite(title) {
  // Every site starts with SharePoint's own three groups.
  const groups = { 3: { Id: 3, Title: `${title} Owners`, members: [] }, 4: { Id: 4, Title: `${title} Members`, members: [] }, 5: { Id: 5, Title: `${title} Visitors`, members: [] } };
  return { title, lists: {}, groups, nextGroup: 10 };
}

function createState(users) {
  const state = { sites: { [SITE]: newSite('Human Resources') }, users, principals: {}, nextPrincipal: 100, log: [] };
  // The HR site is the default one (tests read state.lists / state.groups).
  Object.defineProperty(state, 'lists', { get: () => state.sites[SITE].lists, set: (v) => (state.sites[SITE].lists = v), enumerable: true });
  Object.defineProperty(state, 'groups', { get: () => state.sites[SITE].groups, enumerable: true });
  return state;
}

/** The site a request is for (/sites/<name>, or '' for the tenant root site); created empty on first use. */
function siteOf(state, path) {
  const sitePath = path.substring(0, path.indexOf('/_api/'));
  if (!state.sites[sitePath]) state.sites[sitePath] = newSite(sitePath ? sitePath.split('/').pop() : 'Contoso');
  return { sitePath, site: state.sites[sitePath] };
}

function odataError(status, message) {
  return { status, body: { 'odata.error': { code: '-1, Microsoft.SharePoint.SPException', message: { lang: 'en-US', value: message } } } };
}

function parseFieldXml(xml) {
  const attr = (n) => {
    const m = new RegExp(`\\s${n}="([^"]*)"`).exec(xml);
    return m ? m[1] : undefined;
  };
  const choices = [];
  const re = /<CHOICE>([^<]*)<\/CHOICE>/g;
  let m;
  while ((m = re.exec(xml))) choices.push(m[1]);
  return { name: attr('Name'), type: attr('Type'), indexed: attr('Indexed') === 'TRUE', choices, maxLength: Number(attr('MaxLength') || 255) };
}

// OData $filter: eq/ne/lt/le/gt/ge, and/or, parentheses, strings, numbers, booleans, null.
function parseFilter(src, hasField) {
  if (!src) return () => true;
  const tokens = [];
  const re = /\s*(\(|\)|'(?:[^']|'')*'|[\w/.:-]+)/gy;
  let m;
  while ((m = re.exec(src))) tokens.push(m[1]);
  let i = 0;
  const literal = (t) => {
    if (/^'/.test(t)) return t.slice(1, -1).replace(/''/g, "'");
    if (t === 'true' || t === 'false') return t === 'true';
    if (t === 'null') return null;
    return /^-?\d+(\.\d+)?$/.test(t) ? Number(t) : t;
  };
  const primary = () => {
    if (tokens[i] === '(') {
      i++;
      const e = orExpr();
      i++;
      return e;
    }
    const field = tokens[i++];
    const op = (tokens[i++] || '').toLowerCase();
    const value = literal(tokens[i++] || '');
    if (!hasField(field)) throw Object.assign(new Error(`Column '${field}' does not exist. It may have been deleted by another user.`), { status: 400 });
    return (row) => {
      const v = row[field];
      if (value === null) return op === 'eq' ? v == null || v === '' : !(v == null || v === '');
      if (op === 'eq') return typeof value === 'string' ? String(v || '').toLowerCase() === value.toLowerCase() : v === value;
      if (op === 'ne') return typeof value === 'string' ? String(v || '').toLowerCase() !== value.toLowerCase() : v !== value;
      const a = typeof value === 'number' ? Number(v) : String(v);
      return { lt: a < value, le: a <= value, gt: a > value, ge: a >= value }[op];
    };
  };
  const andExpr = () => {
    const parts = [primary()];
    while ((tokens[i] || '').toLowerCase() === 'and') {
      i++;
      parts.push(primary());
    }
    return (row) => parts.every((t) => t(row));
  };
  const orExpr = () => {
    const parts = [andExpr()];
    while ((tokens[i] || '').toLowerCase() === 'or') {
      i++;
      parts.push(andExpr());
    }
    return (row) => parts.some((t) => t(row));
  };
  return orExpr();
}

function query(url) {
  const out = {};
  const i = url.indexOf('?');
  if (i < 0) return out;
  url
    .substring(i + 1)
    .split('&')
    .forEach((pair) => {
      const j = pair.indexOf('=');
      if (j > 0) out[decodeURIComponent(pair.substring(0, j))] = decodeURIComponent(pair.substring(j + 1).replace(/\+/g, ' '));
    });
  return out;
}

function validate(list, data) {
  for (const k of Object.keys(data)) {
    if (k === '__metadata') continue;
    const type = list.fields[k] ? list.fields[k].type : BUILT_IN[k];
    if (!type) return `The property '${k}' does not exist on type 'SP.Data.${list.title.replace(/_/g, '_x005f_')}ListItem'. Make sure to only use property names that are defined by the type.`;
    const v = data[k];
    if (v === null || v === undefined) continue;
    if (type === 'Number' && typeof v !== 'number') return `Cannot convert a primitive value to the expected type 'Edm.Double' (${k}).`;
    if (type === 'Boolean' && typeof v !== 'boolean') return `Cannot convert a primitive value to the expected type 'Edm.Boolean' (${k}).`;
    if (type === 'DateTime' && isNaN(Date.parse(v))) return `Cannot convert '${v}' to a date (${k}).`;
    if (type === 'Text' && String(v).length > (list.fields[k] ? list.fields[k].maxLength : 255)) return `The value of the field '${k}' is too long: ${String(v).length} characters, the limit is 255.`;
  }
  return '';
}

function handle(state, method, url, headers, body, origin) {
  const userEmail = (headers['x-user'] || '').toLowerCase();
  const user = state.users.filter((u) => u.Email.toLowerCase() === userEmail)[0];
  if (!user) return odataError(401, 'Access denied: unknown user ' + userEmail);
  const verb = (headers['x-http-method'] || method).toUpperCase();
  const p = decodeURIComponent(url.split('?')[0]);
  const api = p.substring(p.indexOf('/_api/') + 6);
  const { sitePath, site } = siteOf(state, p);
  const q = query(url);
  const json = () => (body && body.length ? JSON.parse(body.toString('utf8')) : {});
  state.log.push(`${user.Email} ${verb} ${sitePath}/_api/${api}`);

  if (api === 'web' || api === 'web/') {
    // Full permission mask for site owners, read/contribute for everyone else.
    const perms = user.IsSiteAdmin ? { High: '2147483647', Low: '4294967295' } : { High: '432', Low: '1011030767' };
    return { status: 200, body: { Id: 'web-' + sitePath, Title: site.title, EffectiveBasePermissions: perms } };
  }
  // SharePoint search: lists whose title matches the first word of the query, on every site (all readable here).
  if (api === 'search/query') {
    const word = ((q.querytext || '').replace(/^'|'$/g, '').split(/\s+/)[0] || '').toLowerCase();
    const rows = [];
    Object.keys(state.sites).forEach((sp) => {
      Object.keys(state.sites[sp].lists).forEach((t) => {
        if (t.toLowerCase() !== word) return;
        rows.push({ Cells: [{ Key: 'Title', Value: t }, { Key: 'Path', Value: `${origin}${sp}/Lists/${t}/AllItems.aspx` }, { Key: 'SPWebUrl', Value: `${origin}${sp}` }] });
      });
    });
    return { status: 200, body: { PrimaryQueryResult: { RelevantResults: { RowCount: rows.length, Table: { Rows: rows } } } } };
  }
  if (api.startsWith('web/currentuser')) return { status: 200, body: { Id: user.Id, Title: user.Title, Email: user.Email, LoginName: 'i:0#.f|membership|' + user.Email, IsSiteAdmin: !!user.IsSiteAdmin } };
  if (api === 'web/ensureuser') {
    const login = json().logonName;
    if (!state.principals[login]) state.principals[login] = state.nextPrincipal++;
    return { status: 200, body: { Id: state.principals[login] } };
  }
  if (api.startsWith('web/roledefinitions')) return { status: 200, body: { value: [{ Id: 1073741826, RoleTypeKind: 2 }, { Id: 1073741827, RoleTypeKind: 3 }, { Id: 1073741830, RoleTypeKind: 6 }, { Id: 1073741829, RoleTypeKind: 5 }] } };
  if (api === 'web/lists' && method === 'POST') {
    const d = json();
    if (site.lists[d.Title]) return odataError(400, `A list, survey, discussion board, or document library with the specified title already exists in this Web site.  Please choose another title.`);
    if (!user.IsSiteAdmin) return odataError(403, 'Access denied. You do not have permission to perform this action or access this resource.');
    site.lists[d.Title] = { title: d.Title, id: crypto.randomUUID(), created: new Date().toISOString(), fields: {}, items: [], nextId: 1, props: { EnableAttachments: true, Hidden: false }, unique: false, assignments: [] };
    return { status: 201, body: { Id: site.lists[d.Title].id, Title: d.Title } };
  }
  // Site groups: SharePoint's own (Owners 3, Members 4, Visitors 5) and the ones the app creates.
  if (api.startsWith('web/associatedmembergroup')) return { status: 200, body: { Id: 4 } };
  if (api.startsWith('web/associatedvisitorgroup')) return { status: 200, body: { Id: 5 } };
  if (api.startsWith('web/associatedownergroup')) return { status: 200, body: { Id: 3 } };
  const gByName = /^web\/sitegroups\/getbyname\('((?:[^']|'')+)'\)/.exec(api);
  if (gByName) {
    const name = gByName[1].replace(/''/g, "'");
    const g = Object.values(site.groups).filter((x) => x.Title === name)[0];
    return g ? { status: 200, body: { Id: g.Id, Title: g.Title } } : odataError(404, 'Group cannot be found.');
  }
  if (api === 'web/sitegroups' && method === 'POST') {
    if (!user.IsSiteAdmin) return odataError(403, 'Access denied. You do not have permission to perform this action or access this resource.');
    const d = json();
    if (Object.values(site.groups).some((x) => x.Title === d.Title)) return odataError(400, 'The specified name is already in use.');
    const g = { Id: site.nextGroup++, Title: d.Title, Description: d.Description, members: [] };
    site.groups[g.Id] = g;
    return { status: 201, body: { Id: g.Id, Title: g.Title } };
  }
  const gUsers = /^web\/sitegroups\((\d+)\)\/users(?:\/removebyid\((\d+)\))?/.exec(api);
  if (gUsers) {
    const g = site.groups[Number(gUsers[1])];
    if (!g) return odataError(404, 'Group cannot be found.');
    if (gUsers[2]) {
      if (!user.IsSiteAdmin) return odataError(403, 'Access denied.');
      g.members = g.members.filter((m) => m.Id !== Number(gUsers[2]));
      return { status: 200, body: {} };
    }
    if (method === 'POST') {
      if (!user.IsSiteAdmin) return odataError(403, 'Access denied.');
      const login = json().LoginName;
      if (!state.principals[login]) state.principals[login] = state.nextPrincipal++;
      if (!g.members.some((m) => m.LoginName === login)) g.members.push({ Id: state.principals[login], LoginName: login, Email: login.indexOf('membership|') > 0 ? login.split('|').pop() : '' });
      return { status: 201, body: {} };
    }
    return { status: 200, body: { value: g.members } };
  }
  const lm = /^web\/lists\/getbytitle\('((?:[^']|'')+)'\)(.*)$/i.exec(api);
  if (!lm) return odataError(404, 'Not implemented in the emulator: ' + api);
  const title = lm[1].replace(/''/g, "'");
  const rest = lm[2];
  const list = site.lists[title];
  if (!list) return odataError(404, `List '${title}' does not exist at site with URL 'https://contoso.sharepoint.com${sitePath}'.`);
  const hasField = (f) => !!list.fields[f] || !!BUILT_IN[f];

  if (rest === '' || rest === '/') {
    if (verb === 'MERGE') {
      Object.assign(list.props, json());
      return { status: 204 };
    }
    return { status: 200, body: { Id: list.id, Title: list.title, Created: list.created, ItemCount: list.items.length, ...list.props } };
  }
  if (rest.startsWith('/fields/createfieldasxml')) {
    const f = parseFieldXml(json().parameters.SchemaXml);
    if (list.fields[f.name]) return { status: 200, body: { InternalName: f.name + '0' } };
    list.fields[f.name] = f;
    return { status: 201, body: { InternalName: f.name } };
  }
  const fm = /^\/fields\/getbyinternalnameortitle\('([^']+)'\)$/.exec(rest);
  if (fm) {
    if (!list.fields[fm[1]]) return odataError(404, `Column '${fm[1]}' does not exist.`);
    if (verb === 'MERGE') Object.assign(list.fields[fm[1]], { indexed: !!json().Indexed });
    return { status: 204 };
  }
  if (rest.startsWith('/fields')) {
    const value = Object.keys(BUILT_IN).map((n) => ({ InternalName: n, Indexed: false })).concat(Object.keys(list.fields).map((n) => ({ InternalName: n, Indexed: list.fields[n].indexed })));
    return { status: 200, body: { value } };
  }
  if (/^\/breakroleinheritance/.test(rest)) {
    list.unique = true;
    return { status: 200, body: {} };
  }
  if (/^\/roleassignments/.test(rest)) {
    const add = /addroleassignment\(principalid=(\d+),roledefid=(\d+)\)/.exec(rest);
    if (add) list.assignments.push({ principal: Number(add[1]), role: Number(add[2]) });
    const del = /getbyprincipalid\((\d+)\)\/deleteobject/.exec(rest);
    if (del) list.assignments = list.assignments.filter((a) => a.principal !== Number(del[1]));
    return { status: 200, body: {} };
  }
  const im = /^\/items(?:\((\d+)\))?(\/AttachmentFiles(?:\/add\(FileName='((?:[^']|'')+)'\))?)?$/.exec(rest);
  if (!im) return odataError(404, 'Not implemented in the emulator: ' + api);
  const id = im[1] ? Number(im[1]) : undefined;
  const row = id !== undefined ? list.items.filter((x) => x.Id === id)[0] : undefined;
  if (id !== undefined && !row) return odataError(404, 'Item does not exist. It may have been deleted by another user.');
  if (im[2]) {
    if (im[3]) {
      if (!list.props.EnableAttachments) return odataError(400, 'Attachments are not enabled on this list.');
      row.__files = (row.__files || []).concat({ FileName: im[3].replace(/''/g, "'"), size: body ? body.length : 0 });
      row.Attachments = true;
      return { status: 200, body: { FileName: im[3], ServerRelativeUrl: `${sitePath}/Lists/${title}/Attachments/${id}/${im[3]}` } };
    }
    return { status: 200, body: { value: (row.__files || []).map((f) => ({ FileName: f.FileName, ServerRelativeUrl: `${sitePath}/Lists/${title}/Attachments/${id}/${f.FileName}` })) } };
  }
  const pub = (r) => {
    const out = {};
    Object.keys(r).forEach((k) => {
      if (k.indexOf('__') !== 0) out[k] = r[k];
    });
    return out;
  };
  if (method === 'GET' && row) return { status: 200, body: pub(row) };
  if (method === 'GET') {
    let rows;
    try {
      rows = list.items.filter(parseFilter(q.$filter, hasField));
    } catch (e) {
      return odataError(e.status || 400, e.message);
    }
    if (q.$orderby) {
      const [field, dir] = q.$orderby.split(',')[0].trim().split(/\s+/);
      if (!hasField(field)) return odataError(400, `Column '${field}' does not exist.`);
      rows = rows.slice().sort((a, b) => (a[field] < b[field] ? -1 : a[field] > b[field] ? 1 : 0) * (dir === 'desc' ? -1 : 1));
    }
    const top = Math.min(Number(q.$top || 100), 5000);
    const skip = Number(q.$skiptoken || 0);
    const page = rows.slice(skip, skip + top).map(pub);
    const out = { value: page };
    if (skip + top < rows.length) out['odata.nextLink'] = `${origin}${sitePath}/_api/web/lists/getbytitle('${encodeURIComponent(title)}')/items?${url.split('?')[1].replace(/&?\$skiptoken=\d+/, '')}&$skiptoken=${skip + top}`;
    return { status: 200, body: out };
  }
  if (verb === 'DELETE') {
    list.items = list.items.filter((x) => x !== row);
    return { status: 200, body: {} };
  }
  const data = json();
  const problem = validate(list, data);
  if (problem) return odataError(400, problem);
  const now = new Date().toISOString();
  if (verb === 'MERGE') {
    Object.assign(row, data, { Modified: now, EditorId: user.Id });
    return { status: 204 };
  }
  const item = { Id: list.nextId, ID: list.nextId, Title: '', Attachments: false, Created: now, Modified: now, AuthorId: user.Id, EditorId: user.Id, ...data };
  list.nextId++;
  list.items.push(item);
  return { status: 201, body: pub(item) };
}

/** Start the emulator. Serves the e2e bundle at / and the REST API under /sites/hr/_api/. */
function start(port, users, staticDir) {
  let state = createState(users);
  const server = http.createServer((req, res) => {
    const chunks = [];
    req.on('data', (c) => chunks.push(c));
    req.on('end', () => {
      const body = Buffer.concat(chunks);
      const send = (status, payload, type) => {
        res.writeHead(status, { 'Content-Type': type || 'application/json', Date: new Date().toUTCString() });
        res.end(payload === undefined ? '' : typeof payload === 'string' || Buffer.isBuffer(payload) ? payload : JSON.stringify(payload));
      };
      if (req.url === '/__reset' && req.method === 'POST') {
        state = createState(users);
        return send(200, { ok: true });
      }
      if (req.url.startsWith('/__state')) return send(200, state);
      if (req.url.startsWith('/__seed') && req.method === 'POST') {
        // Test set-up shortcut: replace the whole site (lists, columns, items) with the given state.
        state.lists = JSON.parse(body.toString('utf8')).lists;
        return send(200, { ok: true });
      }
      if (req.url.indexOf('/_api/') >= 0) {
        let r;
        try {
          r = handle(state, req.method, req.url, req.headers, body, `http://${req.headers.host}`);
        } catch (e) {
          r = odataError(500, 'Emulator error: ' + e.message);
        }
        return send(r.status, r.body);
      }
      if (req.url.indexOf('/_layouts/') >= 0 && !/teamshostedapp/i.test(req.url)) return send(404, '', 'text/plain');
      const isPage = req.url === '/' || req.url.startsWith('/?') || /\/_layouts\/15\/teamshostedapp\.aspx/i.test(req.url);
      const file = path.join(staticDir, isPage ? 'index.html' : req.url.split('?')[0]);
      if (file.indexOf(staticDir) === 0 && fs.existsSync(file) && fs.statSync(file).isFile()) return send(200, fs.readFileSync(file), file.endsWith('.html') ? 'text/html' : 'application/javascript');
      return send(404, 'not found', 'text/plain');
    });
  });
  return new Promise((resolve) => server.listen(port, () => resolve({ server, getState: () => state })));
}

module.exports = { start, SITE };
