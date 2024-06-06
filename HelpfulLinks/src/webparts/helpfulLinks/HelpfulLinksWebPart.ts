import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneFieldType,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'HelpfulLinksWebPartStrings';
import HelpfulLinks from './components/HelpfulLinks';
import { IHelpfulLinksProps } from './components/IHelpfulLinksProps';

export interface IHelpfulLinksWebPartProps {
  description: string;
  numberOfLinks: number;
  linkNameArr: string[];
  linkUrlArr: string[];
}

export default class HelpfulLinksWebPart extends BaseClientSideWebPart<IHelpfulLinksWebPartProps> {


  private _checkLinkCount(linkCount: number): number {
    let count: number = 1;
    if(linkCount != null)
      count = linkCount;
    return count;
  } 

  public render(): void {
    const element: React.ReactElement<IHelpfulLinksProps> = React.createElement(
      HelpfulLinks,
      {
        description: this.properties.description,
        linkCount: this._checkLinkCount(this.properties.numberOfLinks),
        linkNames: this.properties.linkNameArr,
        linkUrls: this.properties.linkUrlArr,
      }
    );
    
    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    let linkCount: number = this._checkLinkCount(this.properties.numberOfLinks);
    
    let dynamicGroup: any[] = [
      {
        groupName: strings.BasicGroupName,
        groupFields: [
          PropertyPaneTextField('numberOfLinks', {
            label: 'Number Of Helpful Links'
          })
        ]
      }
    ];

    for(let i: number = 0; i < linkCount; i++){
      let iStr: string = i.toString();
      let groupNameStr: string = "Helpful Link " + iStr;
      let linkNameStr: string = "linkNameArr[" + iStr + "]";
      let linkUrlStr: string = "linkUrlArr[" + iStr + "]";
      dynamicGroup.push(
        {
          groupName: groupNameStr,
          groupFields: [
            PropertyPaneTextField(linkNameStr, {
              label: 'Name of Link'
            }),
            PropertyPaneTextField(linkUrlStr, {
              label: 'Url of Link'
            })
          ]
        }
      );
    }
    
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: dynamicGroup
        }
      ]
    };
  }
}
