import { SPClient, SPError, odataString } from './SPClient';
import { IFieldDef, IListDef, LISTS, ListKey, fieldSchemaXml, listTitle } from '../schema/schema';
import { IPerson } from '../models/types';

export type LogLevel = 'info' | 'success' | 'warning' | 'error';
export type ProgressFn = (message: string, level: LogLevel) => void;

interface ISPField {
  InternalName: string;
  Indexed: boolean;
}

// Add the field with its Name as internal name (8) and show it in the default view (16).
const FIELD_OPTIONS = 8 | 16;
// SP.RoleType
const ROLE_READER = 2;
const ROLE_CONTRIBUTOR = 3;

/** Lists everyone may only read; HR and admins edit them through the app. */
const READ_ONLY_FOR_USERS: ListKey[] = ['settings', 'employees', 'audit'];

/** Creates and repairs every list and column of schema/schema.ts. Needs Manage Lists. */
export class Provisioner {
  private readonly _sp: SPClient;
  private readonly _prefix: string;

  public constructor(sp: SPClient, prefix: string) {
    this._sp = sp;
    this._prefix = prefix;
  }

  public async ensureSchema(log: ProgressFn): Promise<{ ok: boolean; errors: string[] }> {
    const errors: string[] = [];
    for (const def of LISTS) {
      try {
        await this._ensureList(def, log);
      } catch (e) {
        const msg = `${listTitle(this._prefix, def.key)}: ${(e as Error).message}`;
        errors.push(msg);
        log(msg, 'error');
      }
    }
    return { ok: errors.length === 0, errors };
  }

  private async _ensureList(def: IListDef, log: ProgressFn): Promise<void> {
    const title = listTitle(this._prefix, def.key);
    const path = this._sp.list(title);
    let exists = true;
    try {
      await this._sp.get(`${path}?$select=Id`);
    } catch (e) {
      if (!(e instanceof SPError && e.status === 404)) throw e;
      exists = false;
    }
    if (!exists) {
      log(`Creating list ${title}`, 'info');
      await this._sp.post('web/lists', { Title: title, Description: def.description, BaseTemplate: 100, AllowContentTypes: false, ContentTypesEnabled: false });
      await this._sp.merge(path, { EnableAttachments: !!def.attachments, Hidden: def.key !== 'requests' });
    }
    const fields = await this._sp.getAll<ISPField>(`${path}/fields?$select=InternalName,Indexed&$top=500`);
    const have: Record<string, ISPField> = {};
    fields.forEach((f) => (have[f.InternalName] = f));
    for (const f of def.fields) {
      if (!have[f.name]) await this._createField(path, title, f, log);
      else if (f.indexed && !have[f.name].Indexed) await this._sp.merge(`${path}/fields/getbyinternalnameortitle(${odataString(f.name)})`, { Indexed: true }).catch(() => undefined);
    }
    log(exists ? `${title} checked` : `${title} created`, 'success');
  }

  private async _createField(path: string, title: string, f: IFieldDef, log: ProgressFn): Promise<void> {
    await this._sp.post(`${path}/fields/createfieldasxml`, { parameters: { SchemaXml: fieldSchemaXml(f), Options: FIELD_OPTIONS } });
    if (f.indexed) await this._sp.merge(`${path}/fields/getbyinternalnameortitle(${odataString(f.name)})`, { Indexed: true }).catch(() => undefined);
    log(`${title}: column ${f.name}`, 'info');
  }

  /**
   * Recommended permissions: the lists stop inheriting from the site; everyone in the organisation may read the
   * settings, employees and audit lists and contribute to the others (requests, ledger, notifications...); HR and the
   * app admins contribute everywhere. Site owners keep Full Control through the owners group.
   */
  public async applySecurity(tenantId: string | undefined, staff: IPerson[], log: ProgressFn): Promise<string[]> {
    const errors: string[] = [];
    const roles = (await this._sp.get<{ value: Array<{ Id: number; RoleTypeKind: number }> }>('web/roledefinitions?$select=Id,RoleTypeKind')).value;
    const roleId = (kind: number): number | undefined => (roles.filter((r) => r.RoleTypeKind === kind)[0] || { Id: undefined }).Id;
    const reader = roleId(ROLE_READER);
    const contributor = roleId(ROLE_CONTRIBUTOR);
    const everyone = tenantId ? `c:0-.f|rolemanager|spo-grid-all-users/${tenantId}` : '';
    let everyoneId: number | undefined;
    if (everyone) {
      try {
        everyoneId = (await this._sp.post<{ Id: number }>('web/ensureuser', { logonName: everyone })).Id;
      } catch (e) {
        errors.push('Everyone except external users: ' + (e as Error).message);
      }
    }
    const staffIds: number[] = [];
    for (const p of staff) {
      try {
        staffIds.push((await this._sp.post<{ Id: number }>('web/ensureuser', { logonName: p.email })).Id);
      } catch (e) {
        errors.push(`${p.email}: ${(e as Error).message}`);
      }
    }
    for (const def of LISTS) {
      const title = listTitle(this._prefix, def.key);
      const path = this._sp.list(title);
      try {
        await this._sp.post(`${path}/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=true)`);
        if (everyoneId !== undefined) {
          await this._sp.post(`${path}/roleassignments/getbyprincipalid(${everyoneId})/deleteobject()`).catch(() => undefined);
          const role = READ_ONLY_FOR_USERS.indexOf(def.key) >= 0 ? reader : contributor;
          if (role) await this._sp.post(`${path}/roleassignments/addroleassignment(principalid=${everyoneId},roledefid=${role})`);
        }
        for (const id of staffIds) {
          if (contributor) await this._sp.post(`${path}/roleassignments/addroleassignment(principalid=${id},roledefid=${contributor})`).catch(() => undefined);
        }
        log(`${title}: permissions applied`, 'success');
      } catch (e) {
        errors.push(`${title}: ${(e as Error).message}`);
        log(`${title}: ${(e as Error).message}`, 'error');
      }
    }
    return errors;
  }
}
