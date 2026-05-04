import * as React from 'react';
import { useState } from 'react';
import { FluentProvider, IdPrefixProvider, MessageBar, MessageBarBody, MessageBarTitle, Rating, Theme } from "@fluentui/react-components";
import styles from './<%= componentName.pascal %>.module.scss';
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
  const [rating, setRating] = useState(5);

  return <IdPrefixProvider value="<%= componentName.camel %>-">
          <FluentProvider theme={props.theme}>
             <WebPartTitle className={styles.<%= componentName.camel %>webPartTitle} displayMode={props.displayMode} title={props.title} updateProperty={props.updateTitleProperty} />

              <MessageBar key="success" intent="success">
                <MessageBarBody>
                  <MessageBarTitle>Congrats!</MessageBarTitle>
                    For creating your first WebPart using the Alight Fluent UI 9 template.
                </MessageBarBody>
              </MessageBar>

              <Rating color="brand" defaultValue={rating} step={0.5} onChange={(_, data) => setRating(data.value)} />
          </FluentProvider>
        </IdPrefixProvider>
}