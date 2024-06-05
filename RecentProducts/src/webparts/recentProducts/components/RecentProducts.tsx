import * as React from 'react';
import styles from './RecentProducts.module.scss';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { IRecentProductsProps } from './IRecentProductsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IDocument } from '../IDocument';
import DocumentComponent from './DocumentComponent';

 // Temp pdf for iframe
 // https://nvlpubs.nist.gov/nistpubs/SpecialPublications/NIST.SP.800-207.pdf
export default class RecentProducts extends React.Component<IRecentProductsProps, {
  showPanel: boolean;
  iframeUrl: string;
}> {
  constructor(props: IRecentProductsProps){
    super(props);
    this.state = { 
      showPanel: false,
      iframeUrl: ""
    };
    this._showPanel = this._showPanel.bind(this);
  } 

  private _showPanel( active: boolean, docUrl: string) { 
    return (): void => {
      this.setState(() => ({
        showPanel : active,
        iframeUrl : docUrl
    }));
  };
  }
  
  public render(): React.ReactElement<IRecentProductsProps> {
    const docs: any[] = [];
    
    this.props.docArr.forEach((doc: IDocument) => {
      const docUrl: string = this.props.docLibUrl + doc.FileLeafRef;
      
      docs.push(
        <div className={styles.column} onClick={ this._showPanel(true,docUrl) }>
          <DocumentComponent
            documentId={doc.Id}
            documentName={doc.Title}
            documentClassification={doc.classification}
            documentDescription={doc.description0}
            documentImgURL={doc.imgUrl}
          ></DocumentComponent>
      </div>
      );
    });
    return (
      <div className={ styles.recentProducts }>
        <div className={styles.row}>
          {docs}
          <div>
            <Panel
              isOpen={ this.state.showPanel }
              onDismiss={ this._setShowPanel(false) }
              type={ PanelType.medium }
              headerText='Document'
            >
              <object width='100%' height='500'>
                <embed width='100%' height='500' type="application/pdf" src={this.state.iframeUrl}></embed>
              </object>
            </Panel>
          </div>
        </div>
      </div>
    );
  }

  

  private _setShowPanel = (showPanel: boolean): () => void => {
    return (): void => {
      this.setState({showPanel});
    };
  }

}
