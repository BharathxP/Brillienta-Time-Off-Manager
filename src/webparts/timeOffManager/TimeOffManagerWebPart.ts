import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneChoiceGroup, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPPermission } from '@microsoft/sp-page-context';
import { SPHttpClient } from '@microsoft/sp-http';
import { initializeIcons } from '@fluentui/react';

import * as strings from 'TimeOffManagerWebPartStrings';
import App, { IAppProps } from './components/App';
import { LIST_PREFIX, listTitle } from './schema/schema';
import { IResolvedSite, TEAMS_CHOICE_KEY, normalizeSiteUrl, resolveSite, sameSite } from './services/site';

export interface ITimeOffManagerWebPartProps {
  /** What the web part shows: the whole app, the team calendar or my leave. */
  view?: 'app' | 'calendar' | 'my';
  /** Site that holds the lists when it is not the page's own site (Teams tabs); empty = found automatically. */
  siteUrl?: string;
}

interface ISiteInfo extends IResolvedSite {
  title: string;
  manageLists: boolean;
  manageWeb: boolean;
}

let iconsReady = false;
const NOMETA = { headers: { Accept: 'application/json;odata=nometadata' } };

export default class TimeOffManagerWebPart extends BaseClientSideWebPart<ITimeOffManagerWebPartProps> {
  private _site: ISiteInfo | undefined;

  private get _inTeams(): boolean {
    return !!(this.context.sdks && this.context.sdks.microsoftTeams);
  }

  protected async onInit(): Promise<void> {
    await super.onInit();
    if (!iconsReady) {
      initializeIcons(undefined, { disableWarnings: true });
      iconsReady = true;
    }
    this._site = await this._resolve();
  }

  private async _getJson<T>(url: string): Promise<T> {
    const res = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1, NOMETA);
    if (!res.ok) throw Object.assign(new Error(`${res.status} ${res.statusText}`), { status: res.status });
    return (await res.json()) as T;
  }

  private _remembered(): string {
    try {
      return window.localStorage.getItem(TEAMS_CHOICE_KEY) || '';
    } catch {
      return '';
    }
  }

  /** The Time Off Manager site (services/site.ts) and what the person may do there. */
  private async _resolve(): Promise<ISiteInfo> {
    const web = this.context.pageContext.web;
    const perms = web.permissions;
    const hostInfo = {
      title: web.title,
      manageLists: perms.hasPermission(SPPermission.manageLists),
      manageWeb: perms.hasPermission(SPPermission.manageWeb) || perms.hasPermission(SPPermission.managePermissions)
    };
    const site = await resolveSite({
      hostUrl: web.absoluteUrl,
      siteSetting: this.properties.siteUrl,
      inTeams: this._inTeams,
      settingsTitle: listTitle(LIST_PREFIX, 'settings'),
      remembered: this._remembered(),
      getJson: (url) => this._getJson(url)
    });
    if (site.error || site.choices || sameSite(site.url, web.absoluteUrl)) return { ...site, ...hostInfo };
    try {
      const w = await this._getJson<{ Title: string; EffectiveBasePermissions: { High: string; Low: string } }>(`${site.url}/_api/web?$select=Title,EffectiveBasePermissions`);
      const p = new SPPermission({ High: Number(w.EffectiveBasePermissions.High), Low: Number(w.EffectiveBasePermissions.Low) });
      return { ...site, title: w.Title, manageLists: p.hasPermission(SPPermission.manageLists), manageWeb: p.hasPermission(SPPermission.manageWeb) || p.hasPermission(SPPermission.managePermissions) };
    } catch (e) {
      const status = (e as { status?: number }).status;
      const why = status === 401 || status === 403 ? 'you do not have access to it' : (e as Error).message;
      return { ...site, ...hostInfo, url: web.absoluteUrl, error: `The Time Off Manager site ${site.url} could not be opened: ${why}.` };
    }
  }

  public render(): void {
    const ctx = this.context;
    const aad = ctx.pageContext.aadInfo;
    const site = this._site as ISiteInfo;
    const props: IAppProps = {
      http: ctx.spHttpClient,
      webUrl: site.url,
      siteTitle: site.title,
      permissions: { manageLists: site.manageLists, manageWeb: site.manageWeb },
      tenantId: aad && aad.tenantId ? aad.tenantId.toString() : undefined,
      prefix: LIST_PREFIX,
      appVersion: this.manifest.version,
      graphFactory: ctx.msGraphClientFactory,
      host: { inTeams: this._inTeams, view: this.properties.view || 'app', siteSource: site.source, siteError: site.error, siteChoices: site.choices }
    };
    ReactDom.render(React.createElement(App, props), this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  /** Apply the site setting only when the pane's Apply button is pressed, not on every keystroke. */
  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected onAfterPropertyPaneChangesApplied(): void {
    this._resolve()
      .then((s) => {
        this._site = s;
        this.render();
      })
      .catch(() => this.render());
  }

  /** Every option lives inside the app (Admin); the pane picks what the web part opens on and, for Teams, the site. */
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const origin = new URL(this.context.pageContext.web.absoluteUrl).origin;
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.ViewGroupName,
              groupFields: [
                PropertyPaneChoiceGroup('view', {
                  label: strings.ViewLabel,
                  options: [
                    { key: 'app', text: strings.ViewApp, checked: !this.properties.view || this.properties.view === 'app' },
                    { key: 'calendar', text: strings.ViewCalendar, checked: this.properties.view === 'calendar' },
                    { key: 'my', text: strings.ViewMy, checked: this.properties.view === 'my' }
                  ]
                })
              ]
            },
            {
              groupName: strings.SiteGroupName,
              groupFields: [
                PropertyPaneTextField('siteUrl', {
                  label: strings.SiteUrlLabel,
                  description: strings.SiteUrlDescription,
                  placeholder: 'https://contoso.sharepoint.com/sites/hr',
                  onGetErrorMessage: (v: string) => (!v || !v.trim() || normalizeSiteUrl(v, origin) ? '' : strings.SiteUrlInvalid),
                  deferredValidationTime: 400
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
