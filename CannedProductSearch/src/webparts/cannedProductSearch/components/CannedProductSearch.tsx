import * as React from 'react';
import styles from './CannedProductSearch.module.scss';
import { ICannedProductSearchProps } from './ICannedProductSearchProps';
import { IProduct } from '../interfaces/IProduct';
import { SPHttpClient } from '@microsoft/sp-http';
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

export default class CannedProductSearch extends React.Component<ICannedProductSearchProps, {
  products: {
    Title: string,
    Intel_x0020_Categories?: any,
    Involved_x0020_Nations?: any,
    publishDate?: string
  }[];

  selectionDetails: string;
}> {
  private _selection: Selection;
  constructor(props: ICannedProductSearchProps){
    super(props);
    this._getProducts = this._getProducts.bind(this);
    this._pushProducts = this._pushProducts.bind(this);
    this.state = {
      products: [
        {
          Title: "",
          Intel_x0020_Categories: "",
          Involved_x0020_Nations: "",
          publishDate: ""
        }
      ],
      selectionDetails: this._getSelectionDetails(),
    };

    this._selection = new Selection({});
  }

  

  public render(): React.ReactElement<ICannedProductSearchProps> {
    const {products, selectionDetails } = this.state;

    const termBoxes: any[] =[];
    for(let i: number = 0; i < this.props.termCount; i++)
      termBoxes.push(
        <div className={styles.column} onClick={this._pushProducts(this.props.terms[i])}>
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
          <div className={styles.row}>
            <DetailsList
              items={this.state.products}
              columns={_columns}
              setKey='set'
              // onActiveItemChanged={}
              layoutMode={DetailsListLayoutMode.fixedColumns}
              compact={true}
              />
          </div>
          <div className={styles.row}>
            { selectionDetails }
          </div>
        </div>
      </div>
    );
  }

  private _getSelectionDetails(): string {
    return (this._selection.getSelection()[0] as any).name;
  }

  private _getProducts(term: string): Promise<IProduct[]> {
    const url: string = this.props.context.pageContext.site.absoluteUrl + "/_api/web/lists/getbytitle('Intelligence')/items?$select=FileLeafRef,Title,publishDate,Intel_x0020_Categories,Involved_x0020_Nations,publishDate&$filter=TaxCatchAll/Term eq '" + term + "'&orderby=Created%20desc";

    return this.props.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then(res=>{
        return res.json();
      })
      .then(json=>{
        return json.value;
      }) as Promise<IProduct[]>;
  }

  private _pushProducts(term: string) {
    return (): void => {
      this._getProducts(term)
        .then(items=>{
          let results: IProduct[] = [];
          items.forEach((item: IProduct)=>{
            results.push({
              Title: item.Title,
              Intel_x0020_Categories: items[0].Intel_x0020_Categories.map(o => o.Label).join(', '),
              Involved_x0020_Nations: item.Involved_x0020_Nations.map(o => o.Label).join(', '),
              publishDate: item.publishDate
            });

          });
          
          this.setState(()=>({
              products: results
          }));
        });
    };
    
  }
}
