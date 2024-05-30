import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'HelloReactWebPartStrings';
import HelloReact from './components/HelloReact';
import { IHelloReactProps } from './components/IHelloReactProps';
import TodoClient from './TodoClient';

export interface IHelloReactWebPartProps {
  description: string;
}

export default class HelloReactWebPart extends BaseClientSideWebPart<IHelloReactWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IHelloReactProps > = React.createElement(
      HelloReact,
      {
        description: this.properties.description,
        todoClient: new TodoClient()
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
