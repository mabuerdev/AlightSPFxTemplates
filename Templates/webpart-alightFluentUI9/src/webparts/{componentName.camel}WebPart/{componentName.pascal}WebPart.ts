import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneLabel,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import packageInfo from '../../../config/package-solution.json';

import * as strings from '<%= componentName.pascal %>WebPartStrings';
import {<%= componentName.pascal %>, I<%= componentName.pascal %>Props} from './components/<%= componentName.pascal %>';

import { createV9Theme } from "@fluentui/react-migration-v8-v9";
import { teamsDarkTheme, teamsLightTheme, Theme, webDarkTheme, webLightTheme } from '@fluentui/react-components';
import { getGraph, getSP } from '../../pnpjsConfig';
import { format } from '../../helpers';

export interface I<%= componentName.pascal %>WebPartProps {
  title: string;
}

export enum AppMode {
  SharePoint, SharePointLocal, Teams, TeamsLocal, Office, OfficeLocal, Outlook, OutlookLocal
}

export default class <%= componentName.pascal %>WebPart extends BaseClientSideWebPart<I<%= componentName.pascal %>WebPartProps> {

  private _isDarkTheme: boolean = false;
  private _appMode: AppMode = AppMode.SharePoint;
  private _theme:Theme = webLightTheme;
  private versionNo:string = "";
  private _sp:SPFI | undefined = undefined;
  private _graph: GraphFI | undefined = undefined;

  public render(): void {
    const element: React.ReactElement<I<%= componentName.pascal %>Props> = React.createElement(
      <%= componentName.pascal %>,
      {
        title: this.properties.title,
          displayMode: this.displayMode,
          updateTitleProperty: (value: string) => {
            this.properties.title = value;
          },
        context: this.context,
        theme: this._appMode === AppMode.Teams || this._appMode === AppMode.TeamsLocal ?
          this._isDarkTheme ? teamsDarkTheme : teamsLightTheme :
          this._appMode === AppMode.SharePoint || this._appMode === AppMode.SharePointLocal ?
            this._isDarkTheme ? webDarkTheme : this._theme :
            this._isDarkTheme ? webDarkTheme : webLightTheme,
        sharePointContext: this._sp,
        graphContext: this._graph
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
       this.versionNo = (packageInfo as any).solution.version;
       const _l = this.context.isServedFromLocalhost;
       if (!!this.context.sdks.microsoftTeams) {
         const teamsContext = await this.context.sdks.microsoftTeams.teamsJs.app.getContext();
         switch (teamsContext.app.host.name.toLowerCase()) {
           case 'teams': this._appMode = _l ? AppMode.TeamsLocal : AppMode.Teams; break;
           case 'office': this._appMode = _l ? AppMode.OfficeLocal : AppMode.Office; break;
           case 'outlook': this._appMode = _l ? AppMode.OutlookLocal : AppMode.Outlook; break;
           default: throw new Error('Unknown host');
         }
       } else this._appMode = _l ? AppMode.SharePointLocal : AppMode.SharePoint;
 
       await super.onInit();
    
      this._sp = getSP(this.context)
      this._graph = getGraph(this.context);
   }

 protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) return;
    this._isDarkTheme = !!currentTheme.isInverted;
    //if the app mode is sharepoint, adjust the fluent ui 9 web light theme to use the sharepoint theme color, teams/dark mode should be fine on default
    if (this._appMode === AppMode.SharePoint || this._appMode === AppMode.SharePointLocal) {
      this._theme = createV9Theme(currentTheme as undefined, webLightTheme);
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            },
            {
                groupName: strings.AlightCredentials,
                groupFields: [
                    PropertyPaneLabel("author", {
                        text: strings.Author,
                    }),
                    PropertyPaneLabel("version", {
                        text: format(
                            strings.VersionText,
                            this.versionNo
                        ),
                    }),
                ],
            },
          ]
        }
      ]
    };
  }
}
