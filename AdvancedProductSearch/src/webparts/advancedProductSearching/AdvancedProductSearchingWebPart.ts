import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'AdvancedProductSearchingWebPartStrings';
import AdvancedProductSearching from './components/AdvancedProductSearching';
import { IAdvancedProductSearchingProps } from './components/IAdvancedProductSearchingProps';

export interface IAdvancedProductSearchingWebPartProps {
  description: string;
}

export default class AdvancedProductSearchingWebPart extends BaseClientSideWebPart<IAdvancedProductSearchingWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IAdvancedProductSearchingProps > = React.createElement(
      AdvancedProductSearching,
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
