import * as React from 'react';
import styles from './DirectoryListing.module.scss';
import { IDirectoryListingProps } from './IDirectoryListingProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IDirItem } from '../IDirItem';
import {
  DetailsList,
  DetailsListLayoutMode,
  Selection
} from 'office-ui-fabric-react/lib/DetailsList';

const _columns = [
  {
    key: 'column1',
    name: 'Site Number',
    fieldName: 'SiteNumber',
    minWidth: 50,
    maxWidth: 200,
    isResizable: true
  },
  {
    key: 'column2',
    name: 'Site Name',
    fieldName: 'SiteName',
    minWidth: 50,
    maxWidth: 200,
    isResizable: true
  },
  {
    key: 'column3',
    name: 'Site Phone',
    fieldName: 'SitePhone',
    minWidth: 50,
    maxWidth: 200,
    isResizable: true
  }
];

export default class DirectoryListing extends React.Component<IDirectoryListingProps, {}> {
  
  
  public render(): React.ReactElement<IDirectoryListingProps> {
    
    const items: {
      SiteName: string,
      SiteNumber: string,
      SitePhone: string
    }[] = [];
    this.props.dirItems.forEach((item: IDirItem)=>{
      items.push({
        SiteName: item.SiteName,
        SiteNumber: item.SiteNumber,
        SitePhone: item.SitePhone
      });
    });
    return (
      <div className={ styles.directoryListing }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <DetailsList
              items={items}
              columns={_columns}
              setKey='set'
              layoutMode={DetailsListLayoutMode.fixedColumns}
              compact={true}
              />
          </div>
        </div>
      </div>
    );
  }
}
