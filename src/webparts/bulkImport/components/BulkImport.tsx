import * as React from 'react';
import styles from './BulkImport.module.scss';
import {
  TextField,
  PrimaryButton,
  MessageBar,
  MessageBarType,
  Stack,
  Spinner,
  SpinnerSize
} from 'office-ui-fabric-react';

import { IBulkImportProps } from './IBulkImportProps';
import { IBulkImportState } from './IBulkImportState';
import { escape } from '@microsoft/sp-lodash-subset';

export default class BulkImport extends React.Component<IBulkImportProps, IBulkImportState> {
  constructor(props: IBulkImportProps, state: IBulkImportState) {
    super(props);

    this.state = {
      listName: '',
      isLoading: false,
      functionResponse: '',
    } 
  }


  private sendImportQueue(): void {
    this.setState({
      isLoading: true,
    })
    console.log(this.state.listName);
    // do a search for the list name
  }

  public render(): React.ReactElement<IBulkImportProps> {
    return (
      <div className={ styles.bulkImport }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <div className={ styles.description }>
                Copy and paste the name of the list into the input below:
              </div>
              <Stack horizontal>
                <TextField
                  label="List Name"
                  onChange={(e,o) => {
                    this.setState({
                      listName: o
                    });
                  }}
                />
                <div className={ styles.buttonHolder }>
                  <PrimaryButton 
                    text="Import"
                    onClick={() => {
                      this.sendImportQueue();
                    }}
                  />   
                </div>
              </Stack>
              <div>
                {this.state.isLoading &&
                  <Spinner ariaLive="polite" label='Finding List' ariaLabel='Finding List' size={SpinnerSize.medium} />
                }
              </div>
              <div>
                {this.state.functionResponse === '200' &&
                  <MessageBar
                    messageBarType={MessageBarType.success}
                  >
                    Import Started
                  </MessageBar>
                }
                {this.state.functionResponse === '500' && 
                  <MessageBar
                    messageBarType={MessageBarType.severeWarning}
                  >
                    There was an error, list not found etc
                  </MessageBar>
                }
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
