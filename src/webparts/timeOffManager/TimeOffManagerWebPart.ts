import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneChoiceGroup } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPPermission } from '@microsoft/sp-page-context';
import { initializeIcons } from '@fluentui/react';

import * as strings from 'TimeOffManagerWebPartStrings';
import App, { IAppProps } from './components/App';
import { LEGACY_LIST_PREFIX, LIST_PREFIX } from './schema/schema';
import { SPHttpClient } from '@microsoft/sp-http';

export interface ITimeOffManagerWebPartProps {
  /** Fixed at "BrillTM_" by the manifest: the lists are found through it. */
  listPrefix: string;
  /** What the web part shows: the whole app, the team calendar or my leave. */
  view?: 'app' | 'calendar' | 'my';
}

let iconsReady = false;

const clean = (p: string | undefined): string => (p || '').replace(/[^A-Za-z0-9_]/g, '');

export default class TimeOffManagerWebPart extends BaseClientSideWebPart<ITimeOffManagerWebPartProps> {
  private _prefix: string = LIST_PREFIX;

  protected async onInit(): Promise<void> {
    await super.onInit();
    this._prefix = await this._resolvePrefix();
    if (!iconsReady) {
      initializeIcons(undefined, { disableWarnings: true });
      iconsReady = true;
    }
  }

  public render(): void {
    const ctx = this.context;
    const perms = ctx.pageContext.web.permissions;
    const aad = ctx.pageContext.aadInfo;
    const props: IAppProps = {
      http: ctx.spHttpClient,
      webUrl: ctx.pageContext.web.absoluteUrl,
      siteTitle: ctx.pageContext.web.title,
      permissions: {
        manageLists: perms.hasPermission(SPPermission.manageLists),
        manageWeb: perms.hasPermission(SPPermission.manageWeb) || perms.hasPermission(SPPermission.managePermissions)
      },
      tenantId: aad && aad.tenantId ? aad.tenantId.toString() : undefined,
      prefix: this._prefix,
      appVersion: this.manifest.version,
      graphFactory: ctx.msGraphClientFactory,
      host: { inTeams: !!(ctx.sdks && ctx.sdks.microsoftTeams), view: this.properties.view || 'app' }
    };
    ReactDom.render(React.createElement(App, props), this.domElement);
  }

  /**
   * The lists' prefix: BrillTM_ (the page's own setting when it names another). A site set up before 1.2 has TO_ lists
   * and no BrillTM_ ones: it keeps using them.
   */
  private async _resolvePrefix(): Promise<string> {
    const set = clean(this.properties.listPrefix);
    if (set && set !== LIST_PREFIX && set !== LEGACY_LIST_PREFIX) return set;
    const exists = async (prefix: string): Promise<boolean> => {
      try {
        const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${prefix}Settings')?$select=Id`;
        const res = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
        return res.ok;
      } catch {
        return false;
      }
    };
    if (await exists(LIST_PREFIX)) return LIST_PREFIX;
    if (await exists(LEGACY_LIST_PREFIX)) return LEGACY_LIST_PREFIX;
    return LIST_PREFIX;
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  /** Every option lives inside the app (Admin); the pane only picks what the web part opens on. */
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
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
            }
          ]
        }
      ]
    };
  }
}
