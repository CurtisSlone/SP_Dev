export interface IGetChildTermsInTermWithPagingRequest {
    sspId: string; // guid of term store
    lcid: number;
    guid: string; // guid of term
    termsetId: string; // guid of term set
    includeDeprecated: boolean;
    pageLimit: number;
    pagingForward: boolean;
    includeCurrentChild: boolean;
    currentChildId: string;
    webId: string;
    listId: string;
  }