import { HttpClientResponse, SPHttpClient } from '@microsoft/sp-http';
import {
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode
} from 'office-ui-fabric-react/lib/DetailsList';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';

import * as React from 'react';
import { IProduct } from './IProduct';
import { IProductSearchProps } from './IProductSearchProps';
import styles from './ProductSearch.module.scss';



export interface IFindTermSetRequest {
  searchTerms: string;
  lcid: number;
}

export interface IFindTermSetResult {
  Error: string;
  Lm: number;
  Content: any[];
}

export interface IGetChildTermsInTermSetWithPagingRequest {
  sspId: string; // guid of term store
  lcid: number;
  guid: string; // guid of term set
  includeDeprecated: boolean;
  pageLimit: number;
  pagingForward: boolean;
  includeCurrentChild: boolean;
  currentChildId: string;
  webId: string;
  listId: string;
}

export interface IGetGroupsRequest {
  sspId: string; // guid of term store
  webId: string;
  listId: string;
  includeSystemGroup: boolean;
  lcid: number;
}

export interface IGetTermSetsRequest {
  sspId: string; // guid of term store
  guid: string; // guid of term group
  includeNoneTaggableTermset: boolean;
  webId: string;
  listId: string;
  lcid: number;
}

export interface IPickSspsRequest {
  webId: string;
  listId: string;
  lcid: number;
}

export interface ITerm {
  Id: string;
  Label: string;
  Paths: string[];
}

export interface ITermSet {
  Id: string;
  Name: string;
  Owner: string;
}

export interface ITermSetInformation {
  Id: string;
  Nm: string; // Name
  Ow: string; // Owner
  It: boolean; // IsTermSet
}

export const CheckboxList = ({ items, checkboxState, onChange }) => {

  return (
      <div className={styles.row}>
        {items.map((item) => (
          <div key={item.Id} className={styles.column}>
            <div className={styles['checkbox-item']}>
              <input type="checkbox"
                id={item.Id}
                name={item.Label}
                checked={checkboxState[item.Label] || false}
                onChange={onChange} />
              <label htmlFor={item.Id}>{item.Label}</label>
            </div>
          </div>
        ))}
      </div>
  );
};


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

export default class ProductSearch extends React.Component<IProductSearchProps, any> {
  constructor(props: IProductSearchProps) {
    super(props);
    this.state = {
      query: "",
      results: [],
      sspId: "",
      intelCategoriesTerms: [],
      involvedNationsTerms: [],
      showCategories: false,
      showNations: false,
      checkboxState: {},
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


    this.updateQuery = this.updateQuery.bind(this);
    this.findTermSets = this.findTermSets.bind(this);
    this.getGroups = this.getGroups.bind(this);
    this.getTermSets = this.getTermSets.bind(this);
    this.getIntelCategoryTerms = this.getIntelCategoryTerms.bind(this);
    this.getInvolvedNationTerms = this.getInvolvedNationTerms.bind(this);
    this.pickSsps = this.pickSsps.bind(this);
    this._pushProducts = this._pushProducts.bind(this);
    this._handleCheckboxChange = this._handleCheckboxChange.bind(this);
  }

  private _showIntelCategories(active: boolean){
    return (): void => {
      
      this.setState(()=>({
        showCategories: active
      }));
    };
  }

  private _showInvolvedNations(active: boolean){
    return (): void => {
      this.setState(()=>({
        showNations: active
      }));
    };
  }

  private _handleCheckboxChange(event: React.ChangeEvent<HTMLInputElement>) {
    const { name, checked } = event.target;
    
    this.setState(prevState => ({
      checkboxState: {
        ...prevState.checkboxState,
        [name]: checked,
      },
    }));
  }

  public componentDidMount(): void {
    this.setState({ query: "" });
    this.pickSsps();
  }


  private _customSearch(): Promise<IProduct[]>{
    let url: string = this.props.context.pageContext.site.absoluteUrl + "/_api/web/lists/getbytitle('" + this.props.queryList +"')/items?$select=FileLeafRef,Title,Intel_x0020_Categories,Involved_x0020_Nations,publishDate&$filter=";

    const terms: string[] = [];
    for (const termId in this.state.checkboxState) {
      if (this.state.checkboxState[termId]) {
        terms.push(`(TaxCatchAll/Term eq '${termId}')`);
      }
    }

    if (terms.length > 0) {
      url += terms.join(' and ');
    }
    alert(url);
  
    return this.props.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then(res=>{
        return res.json();
      })
      .then(json=>{
        return json.value;
      }) as Promise<IProduct[]>;
  }

  private _pushProducts(){
    return (): void =>{
      this._customSearch()
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
    }
  }

  private _onItemInvoked(item: any): void {
    this.setState(() => ({
      showPanel : true,
      embedUrl : `${this.props.context.pageContext.site.absoluteUrl}/${this.props.docLib}/${item.FileLeafRef}`
  }));
  }

  private _setShowPanel = (showPanel: boolean): () => void => {
    return (): void => {
      this.setState({showPanel});
    };
  }


  public render(): React.ReactElement<IProductSearchProps> {
    let renderResult = "";
    if (this.state.results) {
      renderResult = JSON.stringify(this.state.results, null, 2);
    }
    
    return (
      <div className={styles.productSearch}>
        <div className={`ms-Grid ${styles.container}`}>
          <div className={styles.row}>
                <h2>Intel Categories</h2>
                <button onClick={this._showIntelCategories(!this.state.showCategories)}> Show Categories </button>
                <div className={`${this.state.showCategories ? styles.visible : styles.hidden }`}>
                  <CheckboxList items={this.state.intelCategoriesTerms}
                  checkboxState={this.state.checkboxState}
                  onChange={this._handleCheckboxChange} />
                </div>
                <h2>Involved Nations</h2>
                <button onClick={this._showInvolvedNations(!this.state.showNations)}> Show Categories </button>
                <div className={`${this.state.showNations ? styles.visible : styles.hidden }`}>
                  <CheckboxList items={this.state.involvedNationsTerms}
                  checkboxState={this.state.checkboxState}
                  onChange={this._handleCheckboxChange} />
                </div>
            
          </div>
          <div className={styles.row}>
          <button onClick={this._pushProducts()}>Search</button>
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

  private updateQuery(query: string) {
    this.setState({
      query: query
    });
  }


  /**
   * Method to find term sets having the search value, or having a term with the value
   */
  private findTermSets() {
    const url = this.props.context.pageContext.web.serverRelativeUrl + '/_vti_bin/TaxonomyInternalService.json/FindTermSet';
    const query: IFindTermSetRequest = { 'searchTerms': this.state.query, 'lcid': 1033 };

    this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      body: JSON.stringify(query)
    }).then((response: HttpClientResponse) => {
      if (response.ok) {
        response.json().then((result: any) => {
          let termSetResult: IFindTermSetResult = result.d;
          let returnResults: ITermSet[] = [];
          if (termSetResult.Content) {
            returnResults = termSetResult.Content
              .filter((termInfo: any) => termInfo.It)
              .map((termInfo: any) => ({
                Id: termInfo.Id,
                Name: termInfo.Nm,
                Owner: termInfo.Ow
              }));
          }
          this.setState({
            results: returnResults
          });
        });
      } else {
        console.warn(response.statusText);
      }
    });
  }

  /**
   * Method to list all level one terms in a term set
   */
  private getIntelCategoryTerms(termSetId: string) {
    const url = this.props.context.pageContext.web.serverRelativeUrl + '/_vti_bin/TaxonomyInternalService.json/GetChildTermsInTermSetWithPaging';
    const query: IGetChildTermsInTermSetWithPagingRequest = {
      sspId: this.state.sspId,
      lcid: 1033,
      guid: termSetId,
      includeDeprecated: false,
      pageLimit: 1000,
      pagingForward: true,
      includeCurrentChild: true,
      currentChildId: "00000000-0000-0000-0000-000000000000",
      webId: this.props.context.pageContext.web.id.toString(),
      listId: "00000000-0000-0000-0000-000000000000"
    };

    this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      body: JSON.stringify(query)
    }).then((response: HttpClientResponse) => {
      if (response.ok) {
        response.json().then((result: any) => {
          let returnResults: ITerm[] = [];
          if (result.d.Content) {
            returnResults = result.d.Content.map((term: any) => ({
              Id: term.Id,
              Label: term.Nm,
              Paths: term.Paths
            }));
          }
          this.setState({
            intelCategoriesTerms: returnResults
          });
        });
      } else {
        console.warn(response.statusText);
      }
    });
  }

  private getInvolvedNationTerms(termSetId: string) {
    const url = this.props.context.pageContext.web.serverRelativeUrl + '/_vti_bin/TaxonomyInternalService.json/GetChildTermsInTermSetWithPaging';
    const query: IGetChildTermsInTermSetWithPagingRequest = {
      sspId: this.state.sspId,
      lcid: 1033,
      guid: termSetId,
      includeDeprecated: false,
      pageLimit: 1000,
      pagingForward: true,
      includeCurrentChild: true,
      currentChildId: "00000000-0000-0000-0000-000000000000",
      webId: this.props.context.pageContext.web.id.toString(),
      listId: "00000000-0000-0000-0000-000000000000"
    };

    this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      body: JSON.stringify(query)
    }).then((response: HttpClientResponse) => {
      if (response.ok) {
        response.json().then((result: any) => {
          let returnResults: ITerm[] = [];
          if (result.d.Content) {
            returnResults = result.d.Content.map((term: any) => ({
              Id: term.Id,
              Label: term.Nm,
              Paths: term.Paths
            }));
          }
          this.setState({
            involvedNationsTerms: returnResults
          });
        });
      } else {
        console.warn(response.statusText);
      }
    });
  }

  /**
   * Method to list all term sets in a term group and match term set names for 'Intel Categories' and 'Involved Nations'
   */
  private getTermSets(termGroupId: string) {
    const url = this.props.context.pageContext.web.serverRelativeUrl + '/_vti_bin/TaxonomyInternalService.json/GetTermSets';
    const query: IGetTermSetsRequest = {
      sspId: this.state.sspId,
      guid: termGroupId,
      webId: this.props.context.pageContext.web.id.toString(),
      listId: "00000000-0000-0000-0000-000000000000",
      includeNoneTaggableTermset: true,
      lcid: 1033
    };

    this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      body: JSON.stringify(query)
    }).then((response: HttpClientResponse) => {
      if (response.ok) {
        response.json().then((result: any) => {
          const termSets: ITermSetInformation[] = result.d.Content;
          let termSetIdToFetch: string[];

          // Find the term set ids for 'Intel Categories' and 'Involved Nations'
          termSets.forEach((termSet) => {
            if (termSet.Nm === 'Intel Categories') {
              this.getIntelCategoryTerms(termSet.Id);
            }

            if (termSet.Nm === 'Involved Nations'){
              this.getInvolvedNationTerms(termSet.Id);
            }
          });
        });
      } else {
        console.warn(response.statusText);
      }
    });
  }

  /**
   * Method to list all term stores 
   */
  private pickSsps() {
    const url = this.props.context.pageContext.web.serverRelativeUrl + '/_vti_bin/TaxonomyInternalService.json/PickSsps';
    const query: IPickSspsRequest = {
      lcid: 1033,
      webId: "00000000-0000-0000-0000-000000000000",
      listId: "00000000-0000-0000-0000-000000000000"
    };

    this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      body: JSON.stringify(query)
    }).then((response: HttpClientResponse) => {
      if (response.ok) {
        response.json().then((result: any) => {
          this.setState({
            results: result.d,
            sspId: result.d.Content[0].Id
          });
          this.getGroups();
        });
      } else {
        console.warn(response.statusText);
      }
    });
  }

  /**
   * Method to list all term groups for a term store
   */
  private getGroups() {
    const url = this.props.context.pageContext.web.serverRelativeUrl + '/_vti_bin/TaxonomyInternalService.json/GetGroups';
    const query: IGetGroupsRequest = {
      sspId: this.state.sspId,
      webId: this.props.context.pageContext.web.id.toString(),
      listId: "00000000-0000-0000-0000-000000000000",
      includeSystemGroup: false,
      lcid: 1033
    };

    this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      body: JSON.stringify(query)
    }).then((response: HttpClientResponse) => {
      if (response.ok) {
        response.json().then((result: any) => {
          this.setState({
            results: result.d
          });
          // Assuming we want to fetch term sets after fetching groups
          this.getTermSets(result.d.Content[0].Id);
        });
      } else {
        console.warn(response.statusText);
      }
    });
  }
       
}
