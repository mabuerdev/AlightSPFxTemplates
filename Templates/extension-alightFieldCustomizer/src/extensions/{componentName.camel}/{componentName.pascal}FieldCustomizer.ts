import * as React from 'react';
import * as ReactDOM from 'react-dom';

import { Log } from '@microsoft/sp-core-library';
import {
  BaseFieldCustomizer,
  type IFieldCustomizerCellEventParameters
} from '@microsoft/sp-listview-extensibility';

import { Theme, webDarkTheme, webLightTheme } from '@fluentui/react-components';
import { createV9Theme } from '@fluentui/react-migration-v8-v9';
import { Theme as ThemeV8 } from '@fluentui/react';

import { ThemeProvider } from "@microsoft/sp-component-base";

import * as strings from '<%= componentName.pascal %>FieldCustomizerStrings';
import { <%= componentName.pascal %>, I<%= componentName.pascal %>Props } from './components/<%= componentName.pascal %>';

/**
 * If your field customizer uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface I<%= componentName.pascal %>FieldCustomizerProperties {
  // This is an example; replace with your own property
  sampleText?: string;
}

export enum AppMode {
  SharePoint, SharePointLocal, Teams, TeamsLocal, Office, OfficeLocal, Outlook, OutlookLocal
}


const LOG_SOURCE: string = '<%= componentName.pascal %>FieldCustomizer';

export default class <%= componentName.pascal %>FieldCustomizer
  extends BaseFieldCustomizer<I<%= componentName.pascal %>FieldCustomizerProperties> {

  private _isDarkTheme: boolean = false;
  private _appMode: AppMode = AppMode.SharePoint;
  private _theme:Theme = webLightTheme;
  private _themeProvider:ThemeProvider | undefined = undefined;

  public onInit(): Promise<void> {
    // Add your custom initialization to this method.  The framework will wait
    // for the returned promise to resolve before firing any BaseFieldCustomizer events.
    Log.info(LOG_SOURCE, 'Activated <%= componentName.pascal %>FieldCustomizer with properties:');
    Log.info(LOG_SOURCE, JSON.stringify(this.properties, undefined, 2));
    Log.info(LOG_SOURCE, `The following string should be equal: "<%= componentName.pascal %>FieldCustomizer" and "${strings.Title}"`);

    this._appMode = AppMode.SharePoint;

    // Consume the new ThemeProvider service
    this._themeProvider = this.context.serviceScope.consume(
      ThemeProvider.serviceKey
    );

    const theme = this._themeProvider.tryGetTheme();
    if (theme) {
      this.onThemeChanged(theme);
    } else {
      console.warn("Could not load theme");
    }

    this._themeProvider.themeChangedEvent.add(
      this,
      this.onThemeChanged
    );
    return Promise.resolve();
  }

   protected onThemeChanged(currentTheme:any): void {
     if (!currentTheme) return;
    this._isDarkTheme = !!currentTheme.isInverted;
    //if the app mode is sharepoint, adjust the fluent ui 9 web light theme to use the sharepoint theme color, teams/dark mode should be fine on default
    if (this._appMode === AppMode.SharePoint || this._appMode === AppMode.SharePointLocal) {
      this._theme = createV9Theme(currentTheme as ThemeV8 ,  webLightTheme);
    }
  }

  public onRenderCell(event: IFieldCustomizerCellEventParameters): void {
    // Use this method to perform your custom cell rendering.
    const text: string = `${this.properties.sampleText}: ${event.fieldValue}`;

    const <%= componentName.camel %>: React.ReactElement<I<%= componentName.pascal %>Props> =
      React.createElement(<%= componentName.pascal %>, { 
        text: text, 
        theme: this._isDarkTheme ? webDarkTheme : this._theme,
      } as I<%= componentName.pascal %>Props);

    ReactDOM.render(<%= componentName.camel %>, event.domElement);
  }

  public onDisposeCell(event: IFieldCustomizerCellEventParameters): void {
    // This method should be used to free any resources that were allocated during rendering.
    // For example, if your onRenderCell() called ReactDOM.render(), then you should
    // call ReactDOM.unmountComponentAtNode() here.
    ReactDOM.unmountComponentAtNode(event.domElement);
    super.onDisposeCell(event);
  }
}
