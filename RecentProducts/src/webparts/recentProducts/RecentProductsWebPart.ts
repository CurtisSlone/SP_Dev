import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { SPHttpClient, ISPHttpClientOptions, SPHttpClientResponse} from '@microsoft/sp-http';
import * as strings from 'RecentProductsWebPartStrings';
import RecentProducts from './components/RecentProducts';
import { IRecentProductsProps } from './components/IRecentProductsProps';
import { IDocument } from './IDocument';

export interface IRecentProductsWebPartProps {
  description: string;
  numberOfDocs: number;
}

export default class RecentProductsWebPart extends BaseClientSideWebPart<IRecentProductsWebPartProps> {
  
  public render(): void {
    const element: React.ReactElement<IRecentProductsProps> = React.createElement(
      RecentProducts,
      {
        description: this.properties.description,
        docArr: []
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

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }
}
