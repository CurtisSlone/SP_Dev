import { Version, Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { ISPList } from './ISPList';
import { ISPListItem } from './ISPListItem';
import MockSharePointClient from './MockSharePointClient';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http'
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SpcontentWebPart.module.scss';
import * as strings from 'SpcontentWebPartStrings';

export interface ISpcontentWebPartProps {
  description: string;
}

export default class SpcontentWebPart extends BaseClientSideWebPart<ISpcontentWebPartProps> {

  public render(): void {
   this.domElement.innerHTML = `
   <div>
    <button type='button' class='ms-button'>
      <span class='ms-Button-label'> Create List</span>
    </button>
  </div>
   `;

  this._CreateSharePointList = this._CreateSharePointList.bind(this);
  const button: HTMLButtonElement = this.domElement.getElementsByTagName("BUTTON")[0] as HTMLButtonElement;
  button.addEventListener("click", this._CreateSharePointList);
  }

  private _getSharePointLists(): Promise<ISPList[]> {
    const url: string = this.context.pageContext.web.absoluteUrl + "/_api/web/lists";
    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then(response=>{
        return response.json();
      })
      .then(json=>{
        return json.value;
      }) as Promise<ISPList[]>;
  }

  private _CreateSharePointList(): void {
    const getListURl: string = this.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetByTitle('My List')";
    this.context.spHttpClient.get(getListURl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if(response.status === 200) {
          alert("List already exists.");
          return;
        }

        if (response.status === 404) {
          const url: string = this.context.pageContext.web.absoluteUrl + "/_api/web.lists";
          const listDefinition : any = {
            "Title": "My List",
            "Description": "My description",
            "AllowContentTypes": true,
            "BaseTemplate": 100,
            "ContentTypesEnabled": true,
          };

          const spHttpClientOptions: ISPHttpClientOptions = {
            "body": JSON.stringify(listDefinition),
          };

          this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
            .then((response: SPHttpClientResponse) => {
              if(response.status === 201){
                alert('List created');
              } else {
                alert("Response status "+ response.status + " - " + response.statusText);
              }
            });
        } else {
          alert("Something went wrong.");
        }
      });
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

  private _getMockData(): Promise<ISPListItem[]> {
    return MockSharePointClient.get("")
      .then((data: ISPListItem[]) => {
        return data;
      });
  }

  private _getListItems(): Promise<ISPListItem[]> {
    if (Environment.type === EnvironmentType.Local) {
      return this._getMockData();
    } else {
      alert("TODO: Implement real thing here");
      return null;
    }
  }
}
