import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';
import DirectoryListing from './components/DirectoryListing';
import { IDirectoryListingProps } from './components/IDirectoryListingProps';
import { IDirItem } from './IDirItem';

export interface IDirectoryListingWebPartProps {
  description: string;
  dirItems: IDirItem[];
}

export default class DirectoryListingWebPart extends BaseClientSideWebPart<IDirectoryListingWebPartProps> {

  private _getDirectoryItems(): Promise<IDirItem[]>{
    const url: string = this.context.pageContext.site.absoluteUrl + "/_api/web/lists/getbytitle('directory')/items?$select=SiteName,SiteNumber,SitePhone";

    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then(res=>{
        return res.json();
      })
      .then(json=>{
        return json.value;
      }) as Promise<IDirItem[]>;
  }

  private _pushDirItems(): void {
    this._getDirectoryItems()
      .then(items=>{
        this.properties.dirItems = items;
      });
  }
  public render(): void {
    const element: React.ReactElement<IDirectoryListingProps > = React.createElement(
      DirectoryListing,
      {
        dirItems: this.properties.dirItems,

      }
    );

    this._pushDirItems = this._pushDirItems.bind(this);
    this._pushDirItems();
    
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
