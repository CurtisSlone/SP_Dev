import * as React from 'react';
import styles from './CannedProductSearch.module.scss';
import { ICannedProductSearchProps } from './ICannedProductSearchProps';
import { IProduct } from '../interfaces/IProduct';
import { escape } from '@microsoft/sp-lodash-subset';
import {
  DetailsList,
  DetailsListLayoutMode,
  Selection
} from 'office-ui-fabric-react/lib/DetailsList';
import { DocumentCard,
  DocumentCardType,
} from 'office-ui-fabric-react/lib/DocumentCard';

const _columns = [
  {
    key: 'titleCol',
    name: 'Title',
    fieldName: 'Title',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true
  },
  {
    key: 'categoriesCol',
    name: 'Intel Categories',
    fieldName: 'Intel_x0020_Categories',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true
  },
  {
    key: 'nationsCol',
    name: 'Involved Nations',
    fieldName: 'Involved_x0020_Nations',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true
  },
  {
    key: 'pubDateCol',
    name: 'Publish Date',
    fieldName: 'publishDate',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true
  },

];

export default class CannedProductSearch extends React.Component<ICannedProductSearchProps, {}> {
  constructor(props: ICannedProductSearchProps){
    super(props);
  }

  public render(): React.ReactElement<ICannedProductSearchProps> {
    const termBoxes: any[] =[];
    for(let i: number = 0; i < this.props.termCount; i++)
      termBoxes.push(
        <div className={styles.column} >
          <DocumentCard
            type={DocumentCardType.normal}
            className={styles.docCard}
          >
            <h3 className={styles.largeTitle}>{this.props.termLabels[i]}</h3>
          </DocumentCard>
        </div>
      );
    return (
      <div className={ styles.cannedProductSearch }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            {termBoxes}
          </div>
        </div>
      </div>
    );
  }
}
