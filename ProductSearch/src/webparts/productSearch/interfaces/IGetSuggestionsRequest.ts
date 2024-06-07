export interface IGetSuggestionsRequest {
    start: string; // query
    lcid: number;
    sspList: string; // guid of term store
    termSetList: string; // guid of term set
    anchorId: string;
    isSpanTermStores: boolean; // search in all termstores
    isSpanTermSets: boolean;
    isIncludeUnavailable: boolean;
    isIncludeDeprecated: boolean;
    isAddTerms: boolean;
    isIncludePathData: boolean;
    excludeKeyword: boolean;
    excludedTermset: string;
  }