import * as React from 'react';
import styles from './ProductSearchCards.module.scss';
import { IProductSearchCardsProps } from './IProductSearchCardsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IProduct } from './IProduct';
import { SPHttpClient } from '@microsoft/sp-http';
import {
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  SelectionMode
} from 'office-ui-fabric-react/lib/DetailsList';
import { DocumentCard,
  DocumentCardType,
} from 'office-ui-fabric-react/lib/DocumentCard';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';

const _columns = [
  {
    key: 'titleCol',
    name: 'Title',
    fieldName: 'Title',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true,
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

export default class ProductSearchCards extends React.Component<IProductSearchCardsProps, {
  products: {
    Title: string,
    Intel_x0020_Categories?: any,
    Involved_x0020_Nations?: any,
    publishDate?: string,
    FileLeafRef: string,
    ServerRedirectedEmbedUrl: string
  }[];
  showPanel: boolean;
  embedUrl: string;
}> {

  constructor(props: IProductSearchCardsProps){
    super(props);
    this._getProducts = this._getProducts.bind(this);
    this._pushProducts = this._pushProducts.bind(this);
    this._onItemInvoked = this._onItemInvoked.bind(this);
    this.state = {
      products: [
        {
          Title: "",
          Intel_x0020_Categories: "",
          Involved_x0020_Nations: "",
          publishDate: "",
          FileLeafRef: "",
          ServerRedirectedEmbedUrl: ""
        }
      ],
      showPanel: false,
      embedUrl: ""

    };

    
  }

  public render(): React.ReactElement<IProductSearchCardsProps> {
    
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
      <div className={ styles.productSearchCards }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            {termBoxes}
          </div>
          <div className={styles.row}>
            <DetailsList
              items={this.state.products}
              columns={_columns}
              setKey='set'
              selectionMode={ SelectionMode.none }
              layoutMode={DetailsListLayoutMode.fixedColumns }
              compact={true}
              onItemInvoked={this._onItemInvoked}
              />
          </div>
          <div>
            <Panel
              isOpen={ this.state.showPanel }
              onDismiss={ this._setShowPanel(false) }
              type={ PanelType.medium }
              headerText='Document'
            >
              <object width='100%' height='500'>
                <embed width='100%' height='500' type="application/pdf" src={this.state.embedUrl}></embed>
              </object>
            </Panel>
          </div>
        </div>
      </div>
    );
  }

  private _onItemInvoked(item: any): void {
    this.setState(() => ({
      showPanel : true,
      embedUrl : `${this.props.context.pageContext.site.absoluteUrl}/${item.FileLeafRef}`
  }));
  }


  private _setShowPanel = (showPanel: boolean): () => void => {
    return (): void => {
      this.setState({showPanel});
    };
  }

  private _getProducts(term: string): Promise<IProduct[]> {
    const url: string = this.props.context.pageContext.site.absoluteUrl + "/_api/web/lists/getbytitle('Intelligence')/items?$select=FileLeafRef,Title,PublishDate,Intel_x0020_Categories,Involved_x0020_Nations,PublishDate,ServerRedirectedEmbedUrl&$filter=TaxCatchAll/Term eq '" + term + "'&orderby=Created%20desc";

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
              publishDate: item.publishDate,
              FileLeafRef: item.FileLeafRef,
              ServerRedirectedEmbedUrl: item.ServerRedirectedEmbedUrl
            });

          });
          
          this.setState(()=>({
              products: results
          }));
        });
    };
  }
}
