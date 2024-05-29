import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneLabel,
  PropertyPaneCustomField
} from '@microsoft/sp-webpart-base';

import { escape } from '@microsoft/sp-lodash-subset';

import styles from './HelloWebPaneWebPart.module.scss';
import * as strings from 'HelloWebPaneWebPartStrings';

export interface IHelloWebPaneWebPartProps {
  description: string;
  password: string;
}

export default class HelloWebPaneWebPart extends BaseClientSideWebPart<IHelloWebPaneWebPartProps> {

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.helloWebPane }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.description }">${escape(this.properties.description)}</p>
              <a href="https://aka.ms/spfx" class="${ styles.button }">
                <span class="${ styles.label }">Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>`;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private _customPasswordFieldRender(elem: HTMLElement, context?: any): void {
    if (elem.childElementCount === 0) {
      let label: HTMLLabelElement = document.createElement("label");
      label.className = "ms-Label";
      label.innerText = "Password";
      elem.appendChild(label);
      let br: HTMLBRElement = document.createElement("br");
      elem.appendChild(br);
      let inputElement: HTMLInputElement = document.createElement("input");
      inputElement.type = "password";
      inputElement.name = context;
      this._customPasswordFieldChanged = this._customPasswordFieldChanged.bind(this);
      inputElement.addEventListener("keyup", this._customPasswordFieldChanged);
      inputElement.className = "ms-TextField-field";
      elem.appendChild(inputElement);
    }
  }

  private _customPasswordFieldChanged( event: Event): void {
    let srcElement: HTMLInputElement = event.srcElement as HTMLInputElement;
    this.properties.password = srcElement.value;
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
            },
            {
              groupName: "Custom group",
              groupFields: [
                PropertyPaneCustomField({
                  key: 'password',
                  onRender: (domElement: HTMLElement, context?: any) =>{
                    this._customPasswordFieldRender(domElement, "password");
                  }
                })
              ]
            }
          ]
        },
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: "Custom Group (page 2)",
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            },
            {
              groupName: "Custom group",
              groupFields: [
                PropertyPaneTextField('textboxField', {
                  label: "Enter a custom value"
                }),
                PropertyPaneLabel('labelField', {
                  text: "This is a custom text in PropertyPaneLebel"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
