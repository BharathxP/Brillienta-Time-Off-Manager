// Licence keys and the 14-day trial.
//
// Keys are signed by Brillienta with a private key that never leaves the vendor (tools/license/license.cjs); the app
// holds only the public key, so a key cannot be made up or edited. Each key names the customer's Microsoft 365 tenant
// and the number of users it covers. Format: TO1.<base64url(payload JSON)>.<base64url(ECDSA P-256 signature)>.
//
// The trial runs 14 days from installation (the earliest of the stored install date and the Settings list's creation
// date, which SharePoint sets itself); "now" comes from the SharePoint server, not the PC clock.

export const TRIAL_DAYS = 14;
export const GRACE_DAYS = 14;
const DAY = 86400000;

export interface ILicensePayload {
  v: number;
  id: string;
  product: 'timeoff';
  tenant: string;
  customer: string;
  /** Users covered (base plan users plus purchased add-on users). */
  users: number;
  issued: string;
  expires?: string;
  support?: 'standard' | 'unlimited';
}

export type LicenseState = 'trial' | 'licensed' | 'grace' | 'trialEnded' | 'expired';

export interface ILicenseStatus {
  state: LicenseState;
  /** The app works (trial, licence or grace period). */
  active: boolean;
  /** Users the app may hold as active employees (0 = unlimited). */
  users: number;
  trialEnds: string;
  daysLeft?: number;
  license?: ILicensePayload;
  keyProblem?: string;
}

/**
 * Pricing shown in Admin › Licence and on the request-a-quote email. Change here (vendor) when prices change.
 * Every plan includes all features, free updates and unlimited support.
 */
export const PRICING = {
  baseUsers: 50,
  addOnPackUsers: 50,
  addOnPackUsd: 1,
  extraUserInr: 100,
  contact: 'sales@brillienta.com'
};

/** Add-on packs and the per-user price for a number of users above the base plan. */
export function quote(users: number): { extra: number; packs: number; usd: number; inr: number } {
  const extra = Math.max(0, users - PRICING.baseUsers);
  const packs = Math.ceil(extra / PRICING.addOnPackUsers);
  return { extra, packs, usd: packs * PRICING.addOnPackUsd, inr: extra * PRICING.extraUserInr };
}

function fromB64url(s: string): Uint8Array {
  const b64 = s.replace(/-/g, '+').replace(/_/g, '/') + '==='.slice((s.length + 3) % 4);
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

export function parseLicenseKey(key: string): { payload: ILicensePayload; signed: string; signature: Uint8Array } {
  const clean = (key || '').replace(/\s+/g, '');
  const parts = clean.split('.');
  if (parts.length !== 3 || parts[0] !== 'TO1') throw new Error('This is not a Time Off Manager licence key.');
  let payload: ILicensePayload;
  try {
    payload = JSON.parse(new TextDecoder().decode(fromB64url(parts[1])));
  } catch {
    throw new Error('The licence key is damaged. Copy it again, in full.');
  }
  if (!payload || payload.v !== 1 || payload.product !== 'timeoff' || !payload.tenant || !payload.id) throw new Error('The licence key is not valid for this product.');
  return { payload, signed: parts[0] + '.' + parts[1], signature: fromB64url(parts[2]) };
}

export async function verifySignature(publicKeySpkiB64: string, signed: string, signature: Uint8Array, subtle: SubtleCrypto = crypto.subtle): Promise<boolean> {
  const spki = Uint8Array.from(atob(publicKeySpkiB64), (c) => c.charCodeAt(0));
  const key = await subtle.importKey('spki', spki, { name: 'ECDSA', namedCurve: 'P-256' }, false, ['verify']);
  const data = new Uint8Array(signed.length);
  for (let i = 0; i < signed.length; i++) data[i] = signed.charCodeAt(i);
  return subtle.verify({ name: 'ECDSA', hash: 'SHA-256' }, key, signature, data);
}

export async function checkLicenseKey(key: string, tenantId: string | undefined, publicKey: string, subtle?: SubtleCrypto): Promise<ILicensePayload> {
  const parsed = parseLicenseKey(key);
  let ok = false;
  try {
    ok = await verifySignature(publicKey, parsed.signed, parsed.signature, subtle);
  } catch {
    ok = false;
  }
  if (!ok) throw new Error('The licence key is not genuine (signature check failed).');
  if (!tenantId) throw new Error('This site did not report its Microsoft 365 tenant id, so the licence cannot be checked.');
  if (parsed.payload.tenant.toLowerCase() !== tenantId.toLowerCase()) throw new Error('This licence key was issued for another Microsoft 365 organisation.');
  return parsed.payload;
}

export function licenseStatus(payload: ILicensePayload | undefined, trialStart: Date, now: Date, keyProblem?: string): ILicenseStatus {
  const trialEnd = trialStart.getTime() + TRIAL_DAYS * DAY;
  const daysUntil = (t: number): number => Math.max(0, Math.ceil((t - now.getTime()) / DAY));
  const base = { trialEnds: new Date(trialEnd).toISOString(), keyProblem };
  if (payload) {
    const end = payload.expires ? Date.parse(payload.expires + 'T23:59:59Z') : undefined;
    if (end === undefined || isNaN(end) || now.getTime() <= end) {
      return { ...base, state: 'licensed', active: true, users: payload.users, license: payload, daysLeft: end === undefined || isNaN(end) ? undefined : daysUntil(end) };
    }
    if (now.getTime() <= end + GRACE_DAYS * DAY) return { ...base, state: 'grace', active: true, users: payload.users, license: payload, daysLeft: daysUntil(end + GRACE_DAYS * DAY) };
    return { ...base, state: 'expired', active: false, users: 0, license: payload, daysLeft: 0 };
  }
  if (now.getTime() < trialEnd) return { ...base, state: 'trial', active: true, users: 0, daysLeft: daysUntil(trialEnd) };
  return { ...base, state: 'trialEnded', active: false, users: 0, daysLeft: 0 };
}

export function describeStatus(s: ILicenseStatus): string {
  switch (s.state) {
    case 'trial':
      return `Free trial: ${s.daysLeft} day${s.daysLeft === 1 ? '' : 's'} left`;
    case 'licensed':
      return `Licensed for ${s.users} users${s.license && s.license.expires ? ` until ${s.license.expires}` : ''}`;
    case 'grace':
      return `Licence ended on ${s.license && s.license.expires}; the app stops in ${s.daysLeft} day${s.daysLeft === 1 ? '' : 's'}`;
    case 'trialEnded':
      return 'The 14-day trial has ended';
    default:
      return 'The licence has ended';
  }
}

/** Active employees within the licence (first come, first served by list order); 0 users = unlimited. */
export function licensedEmails(active: Array<{ email: string; id?: number }>, users: number): Set<string> {
  const sorted = active.slice().sort((a, b) => (a.id || 0) - (b.id || 0));
  const allowed = users > 0 ? sorted.slice(0, users) : sorted;
  return new Set(allowed.map((e) => e.email.toLowerCase()));
}
