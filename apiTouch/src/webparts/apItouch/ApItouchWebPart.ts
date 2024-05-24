import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'ApItouchWebPartStrings';
import ApItouch from './components/ApItouch';
import { IApItouchProps } from './components/IApItouchProps';

export interface IApItouchWebPartProps {
  description: string;
}

export default class ApItouchWebPart extends BaseClientSideWebPart<IApItouchWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IApItouchProps > = React.createElement(
      ApItouch,
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
