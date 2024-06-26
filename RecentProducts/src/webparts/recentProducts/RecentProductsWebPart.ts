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
  numberOfDocs: string;
  docArr: IDocument[];
  docLib: string;
  docList: string;
}
 
export default class RecentProductsWebPart extends BaseClientSideWebPart<IRecentProductsWebPartProps> {
  
  private _getDocs(docCount: string): Promise<IDocument[]>{
    let numOfDocs: string = "1";
    if( docCount != null)
      numOfDocs = docCount;
    const url: string = this.context.pageContext.site.absoluteUrl + "/_api/web/lists/getbytitle('" + this.properties.docList + "')/items?$select=Title,Id,classification,description0,imgUrl,PublishDate,FileLeafRef&$orderby=PublishDate desc&$top=" + numOfDocs;

    return this.context.spHttpClient.get(url,SPHttpClient.configurations.v1)
      .then(response=>{
        return response.json();
      })
      .then(json=>{
        return json.value;
      }) as Promise<IDocument[]>;
  }

  private _pushDocs(): void {
    
    this._getDocs(this.properties.numberOfDocs)
      .then(docs=>{this.properties.docArr = docs;});
  }

  public render(): void {
    
    let tmpArr: IDocument[] = this.properties.docArr != null ? this.properties.docArr : [];
    const element: React.ReactElement<IRecentProductsProps> = React.createElement(
      RecentProducts,
      {
        description: this.properties.description,
        docArr: tmpArr,
        docLibUrl: this.context.pageContext.site.absoluteUrl + "/" + this.properties.docLib + "/",
      }
    );
    
    this._pushDocs.bind(this);
    this._pushDocs();
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
                }),
                PropertyPaneTextField('docList', {
                  label: "Name of List To Query"
                }),
                PropertyPaneTextField('docLib', {
                  label: "Document Library"
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
