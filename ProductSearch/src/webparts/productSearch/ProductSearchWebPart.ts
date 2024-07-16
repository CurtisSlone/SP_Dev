import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import * as React from 'react';
import * as ReactDom from 'react-dom';

import * as strings from 'ProductSearchWebPartStrings';
import { IProductSearchProps } from './components/IProductSearchProps';
import ProductSearch from './components/ProductSearch';

export interface IProductSearchWebPartProps {
  description: string;
  listToQuery: string;
  documentLibrary: string;
}

export default class ProductSearchWebPart extends BaseClientSideWebPart<IProductSearchWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IProductSearchProps > = React.createElement(
      ProductSearch,
      {
        context: this.context,
        queryList: this.properties.listToQuery == null ? 'Intelligence' : this.properties.listToQuery,
        docLib: this.properties.documentLibrary == null ? 'Shared Documents' : this.properties.documentLibrary
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
              groupName: "",
              groupFields: [
                PropertyPaneTextField('listToQuery', {
                  label: 'List To Make Queries On'
                })
              ]
            },
            {
              groupName: "",
              groupFields: [
                PropertyPaneTextField('documentLibrary', {
                  label: 'Document Library Where The Products Are Stored'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
