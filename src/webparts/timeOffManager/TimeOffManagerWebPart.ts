import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneChoiceGroup } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPPermission } from '@microsoft/sp-page-context';
import { initializeIcons } from '@fluentui/react';

import * as strings from 'TimeOffManagerWebPartStrings';
import App, { IAppProps } from './components/App';

export interface ITimeOffManagerWebPartProps {
  /** Fixed at "TO_" by the manifest: the lists are found through it. */
  listPrefix: string;
  /** What the web part shows: the whole app, the team calendar or my leave. */
  view?: 'app' | 'calendar' | 'my';
}

let iconsReady = false;

export default class TimeOffManagerWebPart extends BaseClientSideWebPart<ITimeOffManagerWebPartProps> {
  protected async onInit(): Promise<void> {
    await super.onInit();
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
      prefix: (this.properties.listPrefix || 'TO_').replace(/[^A-Za-z0-9_]/g, '') || 'TO_',
      appVersion: this.manifest.version,
      graphFactory: ctx.msGraphClientFactory,
      host: { inTeams: !!(ctx.sdks && ctx.sdks.microsoftTeams), view: this.properties.view || 'app' }
    };
    ReactDom.render(React.createElement(App, props), this.domElement);
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
