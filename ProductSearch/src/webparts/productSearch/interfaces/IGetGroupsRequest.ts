export interface IGetGroupsRequest {
    sspId: string; // guid of term store
    webId: string;
    listId: string;
    includeSystemGroup: boolean;
    lcid: number;
  }