import { SPHttpClient, ISPHttpClientOptions, SPHttpClientResponse } from '@microsoft/sp-http';

export type SPItem = Record<string, unknown>;

export class SPError extends Error {
  public readonly status: number;
  public constructor(message: string, status: number) {
    super(message);
    this.status = status;
    // Keep instanceof working after down-levelling to ES5.
    Object.setPrototypeOf(this, SPError.prototype);
  }
}

const NOMETA = 'application/json;odata=nometadata';

/** Escape a value for use inside a single-quoted OData / REST string literal. */
export function odataString(value: string): string {
  return "'" + value.replace(/'/g, "''") + "'";
}

/**
 * Thin wrapper around SPHttpClient for the site's REST API (_api). All requests are relative to the app's site.
 */
export class SPClient {
  private readonly _http: SPHttpClient;
  public readonly webUrl: string;

  public constructor(http: SPHttpClient, webUrl: string) {
    this._http = http;
    this.webUrl = webUrl.replace(/\/$/, '');
  }

  public api(path: string): string {
    return path.indexOf('http') === 0 ? path : `${this.webUrl}/_api/${path.replace(/^\//, '')}`;
  }

  private async _parse<T>(res: SPHttpClientResponse): Promise<T> {
    if (res.status === 204) return undefined as unknown as T;
    const text = await res.text();
    if (!res.ok) {
      let message = `${res.status} ${res.statusText}`;
      try {
        const body = JSON.parse(text);
        const err = body['odata.error'] || body.error;
        if (err && err.message) message = typeof err.message === 'string' ? err.message : err.message.value;
      } catch {
        // not JSON — keep the status line
      }
      throw new SPError(message, res.status);
    }
    if (!text) return undefined as unknown as T;
    return JSON.parse(text) as T;
  }

  public async get<T = SPItem>(path: string): Promise<T> {
    const res = await this._http.get(this.api(path), SPHttpClient.configurations.v1, {
      headers: { Accept: NOMETA }
    });
    return this._parse<T>(res);
  }

  /** GET a collection, following odata.nextLink until `max` rows are read. */
  public async getAll<T = SPItem>(path: string, max: number = 20000): Promise<T[]> {
    let url: string | undefined = this.api(path);
    const rows: T[] = [];
    while (url && rows.length < max) {
      const page: { value: T[]; 'odata.nextLink'?: string } = await this.get(url);
      rows.push(...(page.value || []));
      url = page['odata.nextLink'];
    }
    return rows;
  }

  public async post<T = SPItem>(path: string, body?: unknown, headers?: Record<string, string>): Promise<T> {
    const options: ISPHttpClientOptions = {
      headers: { Accept: NOMETA, 'Content-Type': NOMETA, ...(headers || {}) }
    };
    if (body !== undefined) {
      options.body = body instanceof ArrayBuffer || body instanceof Blob ? body : JSON.stringify(body);
    }
    const res = await this._http.post(this.api(path), SPHttpClient.configurations.v1, options);
    return this._parse<T>(res);
  }

  /** POST using the verbose OData format (needed for a few calls that require __metadata type hints). */
  public async postVerbose<T = SPItem>(path: string, body: unknown, merge: boolean = false): Promise<T> {
    const headers: Record<string, string> = {
      Accept: NOMETA,
      'Content-Type': 'application/json;odata=verbose',
      'odata-version': ''
    };
    if (merge) {
      headers['X-HTTP-Method'] = 'MERGE';
      headers['IF-MATCH'] = '*';
    }
    const res = await this._http.post(this.api(path), SPHttpClient.configurations.v1, {
      headers,
      body: JSON.stringify(body)
    });
    return this._parse<T>(res);
  }

  /** The SharePoint server's clock (Date header), so licence and trial checks do not depend on the PC clock. */
  public async serverTime(): Promise<Date> {
    try {
      const res = await this._http.get(this.api('web?$select=Id'), SPHttpClient.configurations.v1);
      const d = Date.parse(res.headers.get('Date') || '');
      return isNaN(d) ? new Date() : new Date(d);
    } catch {
      return new Date();
    }
  }

  /** Raw bytes of a file in this site (e.g. a list item attachment), by server-relative URL. */
  public async getFile(serverRelativeUrl: string): Promise<ArrayBuffer> {
    const url = this.api(`web/GetFileByServerRelativePath(decodedurl=@a)/$value?@a=${encodeURIComponent(odataString(serverRelativeUrl))}`);
    const res = await this._http.get(url, SPHttpClient.configurations.v1);
    if (!res.ok) throw new SPError(`Could not read ${serverRelativeUrl} (${res.status})`, res.status);
    return res.arrayBuffer();
  }

  public async merge(path: string, body: unknown): Promise<void> {
    await this.post(path, body, { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' });
  }

  public async del(path: string): Promise<void> {
    await this.post(path, undefined, { 'X-HTTP-Method': 'DELETE', 'IF-MATCH': '*' });
  }

  public list(title: string): string {
    return `web/lists/getbytitle(${odataString(title)})`;
  }

  public async ensureUser(logonName: string): Promise<number> {
    const u = await this.post<{ Id: number }>('web/ensureuser', { logonName });
    return u.Id;
  }
}

/** Run async work over items with at most `limit` in flight. */
export async function mapLimit<T, R>(items: T[], limit: number, fn: (item: T, index: number) => Promise<R>): Promise<R[]> {
  const out: R[] = new Array(items.length);
  let next = 0;
  const worker = async (): Promise<void> => {
    while (next < items.length) {
      const i = next++;
      out[i] = await fn(items[i], i);
    }
  };
  await Promise.all(Array.from({ length: Math.min(limit, items.length) }, () => worker()));
  return out;
}
