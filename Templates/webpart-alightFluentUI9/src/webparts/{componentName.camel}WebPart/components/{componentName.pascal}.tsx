import * as React from 'react';
import { FluentProvider, IdPrefixProvider, Theme } from "@fluentui/react-components";
import styles from './<%= componentName.pascal %>.module.scss';
import type { I<%= componentName.pascal %>Props } from './I<%= componentName.pascal %>Props';
import { WebPartTitle } from "@pnp/spfx-controls-react";
import { DisplayMode } from "@microsoft/sp-core-library";
import { WebPartContext } from "@microsoft/sp-webpart-base";

import { SPFI } from '@pnp/sp';
import { GraphFI } from '@pnp/graph';

export interface I<%= componentName.pascal %>Props {
  title: string;
  context: WebPartContext;
  displayMode:DisplayMode;
  updateTitleProperty:((value:string) => void);
  theme: Theme;
  sharePointContext: SPFI;
  graphContext: GraphFI;
}


export const <%= componentName.pascal %> = (props:I<%= componentName.pascal %>Props):JSX.Element => {

  return <IdPrefixProvider value="<%= componentName.camel %>-">
          <FluentProvider theme={props.theme}>
             <WebPartTitle className={styles.<%= componentName.camel %>webPartTitle} displayMode={props.displayMode} title={props.title} updateProperty={props.updateTitleProperty} />
          </FluentProvider>
        </IdPrefixProvider>
}