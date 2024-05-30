import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import * as strings from 'RecentProductsWebPartStrings';
import RecentProducts from './components/RecentProducts';
import { IRecentProductsProps } from './components/IRecentProductsProps';
import DocumentClient from './DocumentClient';

export interface IRecentProductsWebPartProps {
  description: string;
  numberOfDocs: number;
}

export default class RecentProductsWebPart extends BaseClientSideWebPart<IRecentProductsWebPartProps> {
  
  public render(): void {
    const element: React.ReactElement<IRecentProductsProps > = React.createElement(
      RecentProducts,
      {
        description: this.properties.description,
        docCount: this.properties.numberOfDocs,
        documentClient: new DocumentClient()
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
            description: "Define how many recent products you would like."
          },
          groups: [
            {
              groupName: "",
              groupFields: [
                PropertyPaneTextField('numberOfDocs', {
                  label: "Number of Recent Products"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
