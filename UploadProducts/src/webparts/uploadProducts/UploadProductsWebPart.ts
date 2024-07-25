import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'UploadProductsWebPartStrings';
import UploadProducts from './components/UploadProducts';
import { IUploadProductsProps } from './components/IUploadProductsProps';

export interface IUploadProductsWebPartProps {
  description: string;
}

export default class UploadProductsWebPart extends BaseClientSideWebPart<IUploadProductsWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IUploadProductsProps > = React.createElement(
      UploadProducts,
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
