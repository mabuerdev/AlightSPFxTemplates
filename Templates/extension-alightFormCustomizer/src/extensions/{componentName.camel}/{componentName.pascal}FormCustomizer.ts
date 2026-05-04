import * as React from 'react';
import * as ReactDOM from 'react-dom';

import { Log } from '@microsoft/sp-core-library';
import {
  BaseFormCustomizer
} from '@microsoft/sp-listview-extensibility';

import { getGraph, getSP } from '../../pnpjsConfig';
import { Theme, webDarkTheme, webLightTheme } from '@fluentui/react-components';
import { createV9Theme } from '@fluentui/react-migration-v8-v9';
import { Theme as ThemeV8 } from '@fluentui/react';
import { SPFI } from '@pnp/sp';
import { GraphFI } from '@pnp/graph';
import PnPTelemetry from "@pnp/telemetry-js";

import { ThemeProvider } from "@microsoft/sp-component-base";

import { <%= componentName.pascal %>, I<%= componentName.pascal %>Props } from './components/<%= componentName.pascal %>';

/**
 * If your form customizer uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface I<%= componentName.pascal %>FormCustomizerProperties {
  // This is an example; replace with your own property
  sampleText?: string;
}

export enum AppMode {
  SharePoint, SharePointLocal, Teams, TeamsLocal, Office, OfficeLocal, Outlook, OutlookLocal
}

const LOG_SOURCE: string = '<%= componentName.pascal %>FormCustomizer';
export default class <%= componentName.pascal %>FormCustomizer
  extends BaseFormCustomizer<I<%= componentName.pascal %>FormCustomizerProperties> {

  private _isDarkTheme: boolean = false;
  private _appMode: AppMode = AppMode.SharePoint;
  private _theme:Theme = webLightTheme;
  private _sp:SPFI | undefined = undefined;
  private _graph: GraphFI | undefined = undefined;

  private _themeProvider:ThemeProvider | undefined = undefined;

  public onInit(): Promise<void> {
    // Add your custom initialization to this method. The framework will wait
    // for the returned promise to resolve before rendering the form.
    Log.info(LOG_SOURCE, 'Activated <%= componentName.pascal %>FormCustomizer with properties:');
    Log.info(LOG_SOURCE, JSON.stringify(this.properties, undefined, 2));
    this._appMode = AppMode.SharePoint;

    const telemetry = PnPTelemetry.getInstance();
    telemetry.optOut();

    this._sp = getSP(this.context)
    this._graph = getGraph(this.context);

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

  public render(): void {
    // Use this method to perform your custom rendering.
    const <%= componentName.camel %>: React.ReactElement<I<%= componentName.pascal %>Props> =
      React.createElement(<%= componentName.pascal %>, {
        context: this.context,
        displayMode: this.displayMode,
        onSave: this._onSave,
        onClose: this._onClose,
        theme: this._isDarkTheme ? webDarkTheme : this._theme,
        sharePointContext: this._sp,
        graphContext: this._graph
       } as I<%= componentName.pascal %>Props);

    ReactDOM.render(<%= componentName.camel %>, this.domElement);
  }

   protected onThemeChanged(currentTheme:any): void {
     if (!currentTheme) return;
    this._isDarkTheme = !!currentTheme.isInverted;
    //if the app mode is sharepoint, adjust the fluent ui 9 web light theme to use the sharepoint theme color, teams/dark mode should be fine on default
    if (this._appMode === AppMode.SharePoint || this._appMode === AppMode.SharePointLocal) {
      this._theme = createV9Theme(currentTheme as ThemeV8 ,  webLightTheme);
    }
  }

  public onDispose(): void {
    // This method should be used to free any resources that were allocated during rendering.
    ReactDOM.unmountComponentAtNode(this.domElement);
    super.onDispose();
  }

  private _onSave = (): void => {
    // You MUST call this.formSaved() after you save the form.
    this.formSaved();
  }

  private _onClose =  (): void => {
    // You MUST call this.formClosed() after you close the form.
    this.formClosed();
  }
}
