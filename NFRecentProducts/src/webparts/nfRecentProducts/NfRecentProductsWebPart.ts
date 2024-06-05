import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient, ISPHttpClientOptions, SPHttpClientResponse} from '@microsoft/sp-http';
import { IDocument } from './IDocument';
import styles from './NfRecentProductsWebPart.module.scss';
import * as strings from 'NfRecentProductsWebPartStrings';

export interface INfRecentProductsWebPartProps {
  description: string;
  docCount: string;
}

export default class NfRecentProductsWebPart extends BaseClientSideWebPart<INfRecentProductsWebPartProps> {

  private _docList: HTMLUListElement = null;

  private _getDocs(docCount: string): Promise<IDocument[]>{

    const url: string = this.context.pageContext.site.absoluteUrl + "/_api/web/lists/getbytitle('Intelligence')/items?$select=Title,Id,classification,description0,imgUrl,publishDate&$orderby=publishDate desc&$top=" + docCount;

    return this.context.spHttpClient.get(url,SPHttpClient.configurations.v1)
      .then(response=>{
        return response.json();
      })
      .then(json=>{
        return json.value;
      }) as Promise<IDocument[]>;
  }

  private _renderDocs(): void {
    this._getDocs(this.properties.docCount)
      .then(docs=>{
        let docString: string = "";
        docs.forEach(doc=>{
          docString += `<li>`;
          docString += `<div>${doc.Id}</div>`;
          docString += `<div>${doc.Title}</div>`;
          docString += `<div>${doc.classification}</div>`;
          docString += `<div>${doc.description}</div>`;
          docString += `<div>${doc.imgUrl}</div>`;
          docString += `</li>`;
        });
        this._docList.innerHTML = docString;
      });
  }
  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.nfRecentProducts }">
        <div class="${ styles.container }">
          <ul>
          </ul>
        </div>
      </div>`;

      this._docList = this.domElement.getElementsByTagName("UL")[0] as HTMLUListElement;
      this._renderDocs = this._renderDocs.bind(this);
      this._renderDocs();
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
                PropertyPaneTextField('docCount', {
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
