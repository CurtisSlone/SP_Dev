import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IProductSearchCardsProps {
  context: WebPartContext;
  termCount: number;
  termLabels: string[];
  terms: string[];
  queryList: string;
  docLib: string;
}
