import { HttpClientResponse, SPHttpClient } from '@microsoft/sp-http';
import * as React from 'react';
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


export default class ProductSearch extends React.Component<IProductSearchProps, any> {
  constructor(props: IProductSearchProps) {
    super(props);
    this.state = {
      query: "",
      results: [],
      sspId: ""
    };

    this.updateQuery = this.updateQuery.bind(this);
    this.findTermSets = this.findTermSets.bind(this);
    this.getGroups = this.getGroups.bind(this);
    this.getTermSets = this.getTermSets.bind(this);
    this.getChildTermsInTermSetWithPaging = this.getChildTermsInTermSetWithPaging.bind(this);
    this.pickSsps = this.pickSsps.bind(this);
  }

  public componentDidMount(): void {
    this.setState({ query: "" });
    this.pickSsps();
  }

  public render(): React.ReactElement<IProductSearchProps> {
    let renderResult = "";
    if (this.state.results) {
      renderResult = JSON.stringify(this.state.results, null, 2);
    }

    return (
      <div className={styles.productSearch}>
        <div className={`ms-Grid ${styles.container}`}>
          <div className="ms-Grid-row">
            <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <pre>
                {renderResult}
              </pre>
            </div>
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
  private getChildTermsInTermSetWithPaging(termSetId: string) {
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
              Label: term.Label,
              Paths: term.Paths
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
          let termSetIdToFetch: string | undefined;

          // Find the term set ids for 'Intel Categories' and 'Involved Nations'
          termSets.forEach((termSet) => {
            if (termSet.Nm === 'Intel Categories' || termSet.Nm === 'Involved Nations') {
              termSetIdToFetch = termSet.Id;
            }
          });

          // If found, fetch child terms in the identified term set
          if (termSetIdToFetch) {
            this.getChildTermsInTermSetWithPaging(termSetIdToFetch);
          }
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

  /**
   * Method to get child terms in a specific term set by name
   * @param termSetName Name of the term set to fetch terms from
   */
  private getTermsInTermSet(termSetName: string) {
    const url = this.props.context.pageContext.web.serverRelativeUrl + '/_vti_bin/TaxonomyInternalService.json/GetTermSets';
    const query: IGetTermSetsRequest = {
      sspId: this.state.sspId,
      guid: "00000000-0000-0000-0000-000000000000", // Replace with the correct term group id if needed
      includeNoneTaggableTermset: true,
      webId: this.props.context.pageContext.web.id.toString(),
      listId: "00000000-0000-0000-0000-000000000000",
      lcid: 1033
    };

    this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      body: JSON.stringify(query)
    }).then((response: HttpClientResponse) => {
      if (response.ok) {
        response.json().then((result: any) => {
          const termSets: ITermSetInformation[] = result.d.Content;
        let termSet: ITermSetInformation | undefined;

        for (let i = 0; i < termSets.length; i++) {
          if (termSets[i].Nm === termSetName) {
            termSet = termSets[i];
            break;
          }
        }

          if (termSet) {
            this.getChildTermsInTermSetWithPaging(termSet.Id);
          } else {
            console.warn(`Term set '${termSetName}' not found.`);
          }
        });
      } else {
        console.warn(response.statusText);
      }
    });
  }
}
