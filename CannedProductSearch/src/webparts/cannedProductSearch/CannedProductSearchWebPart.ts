import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
 
import * as strings from 'CannedProductSearchWebPartStrings';
import CannedProductSearch from './components/CannedProductSearch';
import { ICannedProductSearchProps } from './components/ICannedProductSearchProps';

export interface ICannedProductSearchWebPartProps {
  numberOfTerms: number;
  termBoxLabels: string[];
  termBoxTerms: string[];
  listToQuery: string;
  documentLibrary: string;
}

export default class CannedProductSearchWebPart extends BaseClientSideWebPart<ICannedProductSearchWebPartProps> {

  private _checkTermCount(termCount: number): number {
    let count: number = 1;
    if(termCount != null)
      count = termCount;
    return count;
  }

  public render(): void {
    const element: React.ReactElement<ICannedProductSearchProps > = React.createElement(
      CannedProductSearch,
      {
        context: this.context,
        termCount: this._checkTermCount(this.properties.numberOfTerms),
        termLabels: this.properties.termBoxLabels == null ? [] : this.properties.termBoxLabels,
        terms: this.properties.termBoxTerms == null ? [] : this.properties.termBoxTerms,
        queryList: this.properties.listToQuery == null ? 'Intelligence' : this.properties.listToQuery,
        docLib: this.properties.documentLibrary == null ? 'Shared Documents' : this.properties.documentLibrary
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    let termCount = this._checkTermCount(this.properties.numberOfTerms);
    let dynamicGroup: any[] = [
      {
        groupName: "",
        groupFields: [
          PropertyPaneTextField('listToQuery', {
            label: 'List To Make Queries On'
          })
        ]
      },
      {
        groupName: "",
        groupFields: [
          PropertyPaneTextField('documentLibrary', {
            label: 'Document Library Where The Products Are Stored'
          })
        ]
      },
      {
        groupName: "",
        groupFields: [
          PropertyPaneTextField('numberOfTerms', {
            label: 'Number of Term Searches'
          })
        ]
      }
    ];

    for(let i: number = 0; i < termCount; i++){
      let iStr: string = i.toString();
      let groupNameStr: string = "Term " + iStr;
      let labelStr: string = "termBoxLabels[" + iStr + "]";
      let termStr: string = "termBoxTerms[" + iStr + "]";
      dynamicGroup.push(
        {
          groupName: groupNameStr,
          groupFields: [
            PropertyPaneTextField(labelStr, {
              label: 'Label for Search Box'
            }),
            PropertyPaneTextField(termStr, {
              label: 'Taxonomy Term'
            })
          ]
        }
      );
    }

    return {
      pages: [
        {
          header: {
            description: "Canned Product Searches"
          },
          groups: dynamicGroup
        }
      ]
    };
  }
}
