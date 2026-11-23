import { LEGACY_LIST_PREFIX, LIST_PREFIX } from '../schema/schema';

/** Only letters, digits and underscores (list titles and REST addresses). */
export function cleanPrefix(p: string | undefined): string {
  return (p || '').replace(/[^A-Za-z0-9_]/g, '');
}

/**
 * Which prefix the app's lists use on this site. A page that names its own prefix (anything other than the standard
 * BrillTM_ or the earlier TO_) uses it. Otherwise: BrillTM_ when that site has it; the earlier TO_ lists when only
 * those exist (sites set up before 1.2 keep working); BrillTM_ on a new site.
 * `hasSettings(prefix)` says whether the site has the list <prefix>Settings.
 */
export async function resolveListPrefix(setting: string | undefined, hasSettings: (prefix: string) => Promise<boolean>): Promise<string> {
  const set = cleanPrefix(setting);
  if (set && set !== LIST_PREFIX && set !== LEGACY_LIST_PREFIX) return set;
  if (await hasSettings(LIST_PREFIX)) return LIST_PREFIX;
  if (await hasSettings(LEGACY_LIST_PREFIX)) return LEGACY_LIST_PREFIX;
  return LIST_PREFIX;
}

/** hasSettings over any SPHttpClient-like client (the web part's, or the test harness's). */
export function settingsListProbe(
  webUrl: string,
  get: (url: string) => Promise<{ ok: boolean }>
): (prefix: string) => Promise<boolean> {
  return async (prefix: string) => {
    try {
      const res = await get(`${webUrl.replace(/\/$/, '')}/_api/web/lists/getbytitle('${prefix}Settings')?$select=Id`);
      return res.ok;
    } catch {
      return false;
    }
  };
}
