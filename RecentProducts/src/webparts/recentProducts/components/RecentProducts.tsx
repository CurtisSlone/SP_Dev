import * as React from 'react';
import styles from './RecentProducts.module.scss';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel'
import { IRecentProductsProps } from './IRecentProductsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IDocument } from '../IDocument';
import DocumentComponent from './DocumentComponent';

 // Temp pdf for iframe
 // https://nvlpubs.nist.gov/nistpubs/SpecialPublications/NIST.SP.800-207.pdf
export default class RecentProducts extends React.Component<IRecentProductsProps, {
  showPanel: boolean;
}> {
  constructor(props: IRecentProductsProps){
    super(props);
    this.state = { 
      showPanel: false
    };
  } 
  
  public render(): React.ReactElement<IRecentProductsProps> {
    const docs: any[] = [];
    this.props.docArr.forEach((doc: IDocument) => {
      docs.push(
      <DocumentComponent
        documentId={doc.Id}
        documentName={doc.Title}
        documentClassification={doc.classification}
        documentDescription={doc.description0}
        documentImgURL={doc.imgUrl}
      ></DocumentComponent>
      );
    });
    return (
      <div className={ styles.recentProducts }>
        <div className={styles.row}>
          {docs}
          <div>
            <DefaultButton
              secondaryText='Open Panel'
              onClick={this._setShowPanel(true)}
              text='Open Panel'
            />
            <Panel
              isOpen={ this.state.showPanel }
              onDismiss={ this._setShowPanel(false) }
              type={ PanelType.medium }
              headerText='Document'
            >
              <object>
    <embed type="application/pdf" src="https://nvlpubs.nist.gov/nistpubs/SpecialPublications/NIST.SP.800-207.pdf" ></embed>
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
