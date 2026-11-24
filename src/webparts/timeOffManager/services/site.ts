// Which site holds Time Off Manager's lists, and the page address used in links that leave the browser.
//
// On a SharePoint page the lists live on the page's own site. In Teams the web part runs on another site — the team's
// site for a channel tab, the tenant root site for the personal app — so the Time Off Manager site comes from:
//   1. the tab's "Time Off Manager site" setting, else
//   2. the host site, when it has the lists, else
//   3. SharePoint search for the BrillTM_Settings list (search only returns sites the person can read; no extra
//      permission): one found → it; several → this device's earlier choice, else the one marked "Use this site in
//      Teams", else the person chooses once.
// Pure apart from the injected `getJson`, so the web part and the end-to-end harness share it (tests/teams.test.cjs).

/** Setting (in the site's own Settings list) that marks the site Teams opens when search finds several. */
export const TEAMS_DEFAULT_SETTING = 'TeamsDefault';
/** This device's choice among several sites (localStorage). */
export const TEAMS_CHOICE_KEY = 'brilltm.teamsSite';

/** Pages that host a web part in Teams: links to them only open inside Teams, so they are never shared. */
const TEAMS_HOST_PAGE = /\/_layouts\/15\/(teamshostedapp|teamslogon)\.aspx/i;
/** The local workbench: a page nobody else can open. */
const WORKBENCH_PAGE = /\/_layouts\/15\/workbench\.aspx/i;
/** Parts of a URL below the site itself (pages, lists, libraries, layouts). */
const BELOW_SITE = /\/(SitePages|Pages|Lists|Shared%20Documents|Shared Documents|SiteAssets|_layouts)(\/.*)?$/i;

/**
 * A typed site address (absolute or server-relative, possibly a page or list inside it) as an absolute URL without a
 * trailing slash, query or hash; '' when it is not an address on `origin` (the web part only calls its own tenant).
 */
export function normalizeSiteUrl(input: string, origin: string): string {
  const raw = (input || '').trim();
  if (!raw) return '';
  let url: URL;
  try {
    url = new URL(raw, origin);
  } catch {
    return '';
  }
  if (url.protocol !== 'https:' && url.protocol !== 'http:') return '';
  if (url.origin.toLowerCase() !== origin.replace(/\/$/, '').toLowerCase()) return '';
  return url.origin + url.pathname.replace(BELOW_SITE, '').replace(/\/+$/, '');
}

export function sameSite(a: string, b: string): boolean {
  const n = (u: string): string => (u || '').replace(/\/+$/, '').toLowerCase();
  return !!a && n(a) === n(b);
}

/** A page worth putting in emails: a real SharePoint page, not Teams' host page or the workbench. */
export function isShareablePage(url: string): boolean {
  const u = (url || '').split('#')[0];
  return /^https?:\/\//i.test(u) && !TEAMS_HOST_PAGE.test(u) && !WORKBENCH_PAGE.test(u);
}

/** A page address without its query string and hash (debug flags, tracking parameters, the app's own route). */
export function pageAddress(url: string): string {
  return (url || '').split('#')[0].split('?')[0];
}

/** Base for links in emails and notifications: this page when it is a real page, else the recorded app page. */
export function shareableBase(currentHref: string, storedPageUrl: string): string {
  const current = pageAddress(currentHref);
  if (isShareablePage(current)) return current;
  const stored = pageAddress(storedPageUrl);
  return isShareablePage(stored) ? stored : current;
}

/** SharePoint search query (KQL) for lists titled `settingsTitle`. */
export function siteSearchQuery(settingsTitle: string): string {
  return `${settingsTitle.replace(/[^A-Za-z0-9_]/g, '')} contentclass:STS_List*`;
}

export interface ISearchRow {
  Cells: Array<{ Key: string; Value: string | null }>;
}

/** Site URLs of the settings lists in search results: the list itself (title, or a path in /Lists/<title>), this tenant, each site once. */
export function sitesFromSearch(rows: ISearchRow[], settingsTitle: string, origin: string): string[] {
  const title = settingsTitle.toLowerCase();
  const marker = `/lists/${title}`;
  const out: string[] = [];
  (rows || []).forEach((row) => {
    const cell: Record<string, string> = {};
    (row.Cells || []).forEach((c) => (cell[c.Key] = c.Value || ''));
    let path = cell.Path || '';
    try {
      path = decodeURIComponent(path);
    } catch {
      // keep it encoded
    }
    const at = path.toLowerCase().indexOf(marker);
    const next = at >= 0 ? path.charAt(at + marker.length) : 'x';
    const inList = at >= 0 && (next === '' || next === '/' || next === '?');
    if ((cell.Title || '').toLowerCase() !== title && !inList) return;
    const site = normalizeSiteUrl(cell.SPWebUrl || (inList ? path.slice(0, at) : ''), origin);
    if (site && !out.some((s) => sameSite(s, site))) out.push(site);
  });
  return out;
}

/** Which found site opens: the only one, else this device's choice, else the only marked one; else the person chooses. */
export function pickSite(found: Array<{ url: string; marked: boolean }>, remembered: string): { url: string } | { choose: string[] } | undefined {
  if (!found.length) return undefined;
  if (found.length === 1) return { url: found[0].url };
  const mine = found.filter((f) => sameSite(f.url, remembered))[0];
  if (mine) return { url: mine.url };
  const marked = found.filter((f) => f.marked);
  if (marked.length === 1) return { url: marked[0].url };
  return { choose: (marked.length ? marked : found).map((f) => f.url) };
}

export interface IResolvedSite {
  url: string;
  /** '' = the host site, 'property' = the tab's setting, 'found' = SharePoint search. */
  source: '' | 'property' | 'found';
  /** Several sites found in Teams: the person picks one. */
  choices?: Array<{ url: string; title: string }>;
  error?: string;
}

export interface IResolveOptions {
  hostUrl: string;
  /** The tab's / page's "Time Off Manager site" setting. */
  siteSetting?: string;
  inTeams: boolean;
  /** Title of the settings list, e.g. "BrillTM_Settings". */
  settingsTitle: string;
  /** This device's earlier choice. */
  remembered?: string;
  /** GET JSON from SharePoint REST; rejects on HTTP errors. */
  getJson: <T>(url: string) => Promise<T>;
}

/** Where Time Off Manager's lists are, as described at the top of this file. */
export async function resolveSite(o: IResolveOptions): Promise<IResolvedSite> {
  const origin = new URL(o.hostUrl).origin;
  const host = o.hostUrl.replace(/\/+$/, '');
  const title = o.settingsTitle.replace(/'/g, "''");
  const hasLists = async (url: string): Promise<boolean> => {
    try {
      await o.getJson(`${url}/_api/web/lists/getbytitle('${encodeURIComponent(title)}')?$select=Id`);
      return true;
    } catch {
      return false;
    }
  };
  const typed = (o.siteSetting || '').trim();
  if (typed) {
    const url = normalizeSiteUrl(typed, origin);
    if (!url) return { url: host, source: '', error: `“${typed}” is not a SharePoint site address on this tenant.` };
    return { url, source: sameSite(url, host) ? '' : 'property' };
  }
  if (!o.inTeams || (await hasLists(host))) return { url: host, source: '' };
  const remembered = normalizeSiteUrl(o.remembered || '', origin);
  if (remembered && (await hasLists(remembered))) return { url: remembered, source: 'found' };
  let sites: string[] = [];
  try {
    const q = encodeURIComponent(`'${siteSearchQuery(o.settingsTitle)}'`);
    const r = await o.getJson<{ PrimaryQueryResult?: { RelevantResults?: { Table?: { Rows?: ISearchRow[] } } } }>(
      `${host}/_api/search/query?querytext=${q}&selectproperties='Title,Path,SPWebUrl'&rowlimit=50&trimduplicates=false`
    );
    const rows = (r.PrimaryQueryResult && r.PrimaryQueryResult.RelevantResults && r.PrimaryQueryResult.RelevantResults.Table && r.PrimaryQueryResult.RelevantResults.Table.Rows) || [];
    sites = sitesFromSearch(rows, o.settingsTitle, origin).slice(0, 10);
  } catch {
    return { url: host, source: '' };
  }
  const found = (
    await Promise.all(
      sites.map(async (url) => {
        try {
          const r = await o.getJson<{ value: Array<{ SettingValue: string | null }> }>(
            `${url}/_api/web/lists/getbytitle('${encodeURIComponent(title)}')/items?$select=SettingValue&$filter=Title eq '${TEAMS_DEFAULT_SETTING}'&$top=1`
          );
          return { url, marked: !!r.value && !!r.value[0] && r.value[0].SettingValue === '1' };
        } catch {
          return undefined; // not readable, or not a Time Off Manager site after all
        }
      })
    )
  ).filter((f): f is { url: string; marked: boolean } => !!f);
  const pick = pickSite(found, remembered);
  if (!pick) return { url: host, source: '' };
  if ('url' in pick) return { url: pick.url, source: 'found' };
  const choices = await Promise.all(
    pick.choose.map(async (url) => {
      try {
        return { url, title: (await o.getJson<{ Title: string }>(`${url}/_api/web?$select=Title`)).Title || url };
      } catch {
        return { url, title: url };
      }
    })
  );
  return { url: host, source: '', choices };
}
