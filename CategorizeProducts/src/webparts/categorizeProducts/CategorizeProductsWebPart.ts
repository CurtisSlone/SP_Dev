import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'CategorizeProductsWebPartStrings';
import CategorizeProducts from './components/CategorizeProducts';
import { ICategorizeProductsProps } from './components/ICategorizeProductsProps';

export interface ICategorizeProductsWebPartProps {
  description: string;
}

export default class CategorizeProductsWebPart extends BaseClientSideWebPart<ICategorizeProductsWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ICategorizeProductsProps > = React.createElement(
      CategorizeProducts,
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
