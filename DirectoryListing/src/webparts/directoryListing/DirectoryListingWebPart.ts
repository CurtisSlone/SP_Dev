import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'DirectoryListingWebPartStrings';
import DirectoryListing from './components/DirectoryListing';
import { IDirectoryListingProps } from './components/IDirectoryListingProps';

export interface IDirectoryListingWebPartProps {
  description: string;
}

export default class DirectoryListingWebPart extends BaseClientSideWebPart<IDirectoryListingWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IDirectoryListingProps > = React.createElement(
      DirectoryListing,
      {
        description: this.properties.description
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "Directory List Name"
          },
          groups: [
            {
              groupName: "",
              groupFields: [
                PropertyPaneTextField('description', {
                  label: ""
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
