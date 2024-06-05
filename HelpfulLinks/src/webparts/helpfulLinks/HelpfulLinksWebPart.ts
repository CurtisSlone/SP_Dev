import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'HelpfulLinksWebPartStrings';
import HelpfulLinks from './components/HelpfulLinks';
import { IHelpfulLinksProps } from './components/IHelpfulLinksProps';

export interface IHelpfulLinksWebPartProps {
  description: string;
}

export default class HelpfulLinksWebPart extends BaseClientSideWebPart<IHelpfulLinksWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IHelpfulLinksProps > = React.createElement(
      HelpfulLinks,
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
