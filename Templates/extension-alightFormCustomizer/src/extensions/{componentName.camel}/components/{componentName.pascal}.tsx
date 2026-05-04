import * as React from 'react';
import { useState } from 'react';
import { FormDisplayMode } from '@microsoft/sp-core-library';
import { FormCustomizerContext } from '@microsoft/sp-listview-extensibility';

import { FluentProvider, IdPrefixProvider, MessageBar, MessageBarBody, MessageBarTitle, Rating, Theme } from '@fluentui/react-components';
import { SPFI } from '@pnp/sp';
import { GraphFI } from '@pnp/graph';

import styles from './<%= componentName.pascal %>.module.scss';

export interface I<%= componentName.pascal %>Props {
  context: FormCustomizerContext;
  displayMode: FormDisplayMode;
  onSave: () => void;
  onClose: () => void;
  theme: Theme,
  sharePointContext: SPFI,
  graphContext: GraphFI
}

export const  <%= componentName.pascal %> = (props:I<%= componentName.pascal %>Props): JSX.Element => {
  const [rating, setRating] = useState(5);

  return  <IdPrefixProvider value="<%= componentName.camel %>-">
              <FluentProvider theme={props.theme}>
                <div className={styles.<%= componentName.camel %>}>
                  <MessageBar key="success" intent="success">
                    <MessageBarBody>
                      <MessageBarTitle>Congrats!</MessageBarTitle>
                        For creating your first Forms Customizer using the Alight Fluent UI 9 template.
                    </MessageBarBody>
                  </MessageBar>

                  <Rating color="brand" defaultValue={rating} step={0.5} onChange={(_, data) => setRating(data.value)} />
                </div>
              </FluentProvider>
            </IdPrefixProvider>
}