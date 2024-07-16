import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IProductSearchProps {
  context: WebPartContext;
  queryList: string;
  docLib: string;
  
}
