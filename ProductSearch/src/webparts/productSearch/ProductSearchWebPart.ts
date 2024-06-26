import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'ProductSearchWebPartStrings';
import ProductSearch from './components/ProductSearch';
import { IProductSearchProps } from './components/IProductSearchProps';

export interface IProductSearchWebPartProps {
  description: string;
}

export default class ProductSearchWebPart extends BaseClientSideWebPart<IProductSearchWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IProductSearchProps > = React.createElement(
      ProductSearch,
      {
        context: this.context,
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
            }
          ]
        }
      ]
    };
  }
}
