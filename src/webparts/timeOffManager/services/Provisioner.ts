import { SPClient, SPError, odataString } from './SPClient';
import { GROUP_NAMES, IFieldDef, IListDef, LISTS, ListKey, fieldSchemaXml, listTitle } from '../schema/schema';
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

  // ---------------------------------------------------------------- groups

  private async _group(name: string, description: string, log: ProgressFn): Promise<{ id: number; created: boolean }> {
    try {
      const g = await this._sp.get<{ Id: number }>(`web/sitegroups/getbyname(${odataString(name)})?$select=Id`);
      return { id: g.Id, created: false };
    } catch (e) {
      if (!(e instanceof SPError) || (e.status !== 404 && !/cannot be found|does not exist/i.test(e.message))) throw e;
    }
    const g = await this._sp.postVerbose<{ Id: number }>('web/sitegroups', { __metadata: { type: 'SP.Group' }, Title: name, Description: description });
    log(`Group ${name} created`, 'success');
    return { id: g.Id, created: true };
  }

  private async _members(groupId: number): Promise<Array<{ Id: number; LoginName: string; Email: string }>> {
    const r = await this._sp.get<{ value: Array<{ Id: number; LoginName: string; Email: string }> }>(`web/sitegroups(${groupId})/users?$select=Id,LoginName,Email`);
    return r.value || [];
  }

  private async _addMember(groupId: number, loginName: string): Promise<void> {
    await this._sp.postVerbose(`web/sitegroups(${groupId})/users`, { __metadata: { type: 'SP.User' }, LoginName: loginName });
  }

  /** Make a group hold exactly these people (by email). */
  private async _syncMembers(groupId: number, name: string, people: IPerson[], log: ProgressFn, errors: string[]): Promise<void> {
    const current = await this._members(groupId);
    const want = people.map((p) => p.email.toLowerCase());
    const has = (email: string): boolean => current.some((m) => (m.Email || '').toLowerCase() === email || m.LoginName.toLowerCase().endsWith('|' + email));
    for (const p of people) {
      if (has(p.email.toLowerCase())) continue;
      try {
        await this._addMember(groupId, `i:0#.f|membership|${p.email}`);
        log(`${name}: added ${p.email}`, 'info');
      } catch (e) {
        errors.push(`${name} › ${p.email}: ${(e as Error).message}`);
      }
    }
    for (const m of current) {
      const email = (m.Email || m.LoginName.split('|').pop() || '').toLowerCase();
      if (want.indexOf(email) >= 0) continue;
      try {
        await this._sp.post(`web/sitegroups(${groupId})/users/removebyid(${m.Id})`);
        log(`${name}: removed ${email}`, 'info');
      } catch (e) {
        errors.push(`${name} › remove ${email}: ${(e as Error).message}`);
      }
    }
  }

  /**
   * The app's branded groups (schema GROUP_NAMES): BrillTM Admins and BrillTM HR hold exactly the people under
   * Admin › Access; BrillTM Users starts with "Everyone except external users" and is then left to the admins, so
   * access to the app can be narrowed in SharePoint.
   */
  public async syncGroups(tenantId: string | undefined, admins: IPerson[], hr: IPerson[], log: ProgressFn): Promise<{ ids: { admins: number; hr: number; users: number }; errors: string[] }> {
    const errors: string[] = [];
    const a = await this._group(GROUP_NAMES.admins, 'Brillienta Time Off Manager: app admins (kept in step with Admin › Access).', log);
    const h = await this._group(GROUP_NAMES.hr, 'Brillienta Time Off Manager: HR (kept in step with Admin › Access).', log);
    const u = await this._group(GROUP_NAMES.users, 'Brillienta Time Off Manager: people who may use the app.', log);
    await this._syncMembers(a.id, GROUP_NAMES.admins, admins, log, errors);
    await this._syncMembers(h.id, GROUP_NAMES.hr, hr, log, errors);
    if (u.created) {
      if (tenantId) {
        try {
          await this._addMember(u.id, `c:0-.f|rolemanager|spo-grid-all-users/${tenantId}`);
          log(`${GROUP_NAMES.users}: added Everyone except external users`, 'info');
        } catch (e) {
          errors.push(`${GROUP_NAMES.users} › Everyone except external users: ${(e as Error).message}`);
        }
      } else errors.push(`${GROUP_NAMES.users}: add the people who may use the app (the tenant id is not known here).`);
    }
    return { ids: { admins: a.id, hr: h.id, users: u.id }, errors };
  }

  /** The groups if they exist (nothing is created). */
  public async findGroups(): Promise<Array<{ name: string; id: number; members: number }>> {
    const out: Array<{ name: string; id: number; members: number }> = [];
    for (const name of [GROUP_NAMES.admins, GROUP_NAMES.hr, GROUP_NAMES.users]) {
      try {
        const g = await this._sp.get<{ Id: number }>(`web/sitegroups/getbyname(${odataString(name)})?$select=Id`);
        out.push({ name, id: g.Id, members: (await this._members(g.Id)).length });
      } catch {
        // not created yet
      }
    }
    return out;
  }

  // ---------------------------------------------------------------- list permissions

  /**
   * Recommended permissions. The lists stop inheriting from the site. BrillTM Users may read the settings, employees
   * and audit lists and contribute to the others (requests, ledger, notifications...); BrillTM HR and BrillTM Admins
   * contribute everywhere. The site's own Members and Visitors groups are taken off the app's lists (they would
   * otherwise edit the settings directly); site owners keep Full Control.
   */
  public async applySecurity(tenantId: string | undefined, admins: IPerson[], hr: IPerson[], log: ProgressFn): Promise<string[]> {
    const roles = (await this._sp.get<{ value: Array<{ Id: number; RoleTypeKind: number }> }>('web/roledefinitions?$select=Id,RoleTypeKind')).value;
    const roleId = (kind: number): number | undefined => (roles.filter((r) => r.RoleTypeKind === kind)[0] || { Id: undefined }).Id;
    const reader = roleId(ROLE_READER);
    const contributor = roleId(ROLE_CONTRIBUTOR);
    const groups = await this.syncGroups(tenantId, admins, hr, log);
    const errors = groups.errors.slice();
    const siteGroups: number[] = [];
    for (const which of ['associatedmembergroup', 'associatedvisitorgroup']) {
      try {
        siteGroups.push((await this._sp.get<{ Id: number }>(`web/${which}?$select=Id`)).Id);
      } catch {
        // the site has no such group
      }
    }
    for (const def of LISTS) {
      const title = listTitle(this._prefix, def.key);
      const path = this._sp.list(title);
      const grant = async (principal: number, role: number | undefined): Promise<void> => {
        await this._sp.post(`${path}/roleassignments/getbyprincipalid(${principal})/deleteobject()`).catch(() => undefined);
        if (role) await this._sp.post(`${path}/roleassignments/addroleassignment(principalid=${principal},roledefid=${role})`);
      };
      try {
        await this._sp.post(`${path}/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=true)`);
        for (const id of siteGroups) await this._sp.post(`${path}/roleassignments/getbyprincipalid(${id})/deleteobject()`).catch(() => undefined);
        await grant(groups.ids.users, READ_ONLY_FOR_USERS.indexOf(def.key) >= 0 ? reader : contributor);
        await grant(groups.ids.hr, contributor);
        await grant(groups.ids.admins, contributor);
        log(`${title}: permissions applied`, 'success');
      } catch (e) {
        errors.push(`${title}: ${(e as Error).message}`);
        log(`${title}: ${(e as Error).message}`, 'error');
      }
    }
    return errors;
  }
}
