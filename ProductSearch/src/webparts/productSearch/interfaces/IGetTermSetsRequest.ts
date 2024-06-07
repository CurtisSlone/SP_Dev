export interface IGetTermSetsRequest {
    sspId: string; // guid of term store
    guid: string; // guid of term group
    includeNoneTaggableTermset: boolean;
    webId: string;
    listId: string;
    lcid: number;
  }