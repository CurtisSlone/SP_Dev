import * as React from 'react';
import styles from './RecentProducts.module.scss';
import { IRecentProductsProps } from './IRecentProductsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IDocument } from '../IDocument';
import DocumentComponent from './DocumentComponent';


export default class RecentProducts extends React.Component<IRecentProductsProps, {}> {
  constructor(props: IRecentProductsProps){
    super(props);
  } 
  
  public render(): React.ReactElement<IRecentProductsProps> {
    const docs: any[] = [];
    this.props.docArr.forEach((doc: IDocument) => {
      docs.push(
      <DocumentComponent
        documentId={doc.Id}
        documentName={doc.Title}
        documentClassification={doc.classification}
        documentDescription={doc.description}
        documentImgURL={doc.imgUrl}
        documentUrl='#'
      ></DocumentComponent>
      );
    });
    return (
      <div className={ styles.recentProducts }>
        <div className={styles.row}>
          {docs}
        </div>
      </div>
    );
  }
}
