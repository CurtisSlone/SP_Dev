import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ICannedProductSearchProps {
  context: WebPartContext;
  termCount: number;
  termLabels: string[];
  terms: string[];
}
