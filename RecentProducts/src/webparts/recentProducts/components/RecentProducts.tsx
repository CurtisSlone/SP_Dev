import * as React from 'react';
import styles from './RecentProducts.module.scss';
import { IRecentProductsProps } from './IRecentProductsProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class RecentProducts extends React.Component<IRecentProductsProps, {}> {
  constructor(props: IRecentProductsProps){
    super(props);
    this.state = {
      documentList: []
    };
  }

  private updateDocuments(): void {
    this.props.documentClient.getDocuments(this.props.docCount);
  }

  public render(): React.ReactElement<IRecentProductsProps> {
    return (
      <div className={ styles.recentProducts }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <p className={ styles.description }>{escape(this.props.docCount.toString())}</p>
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
