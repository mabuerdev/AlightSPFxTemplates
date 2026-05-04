import * as React from 'react';
import { useState } from 'react';
import { FluentProvider, IdPrefixProvider, MessageBar, MessageBarBody, MessageBarTitle, Rating, Theme } from '@fluentui/react-components';

import styles from './<%= componentName.pascal %>.module.scss';

export interface I<%= componentName.pascal %>Props {
  text: string;
  theme: Theme,
}

export const  <%= componentName.pascal %> = (props:I<%= componentName.pascal %>Props): JSX.Element => {
  const [rating, setRating] = useState(5);

  return  <IdPrefixProvider value="<%= componentName.camel %>-">
              <FluentProvider theme={props.theme}>
                <div className={styles.<%= componentName.camel %>}>
                  <MessageBar key="success" intent="success">
                    <MessageBarBody>
                      <MessageBarTitle>Congrats!</MessageBarTitle>
                        For creating your first Field Customizer using the Alight Fluent UI 9 template.
                    </MessageBarBody>
                  </MessageBar>

                  <Rating color="brand" defaultValue={rating} step={0.5} onChange={(_, data) => setRating(data.value)} />
                </div>
              </FluentProvider>
            </IdPrefixProvider>
}

