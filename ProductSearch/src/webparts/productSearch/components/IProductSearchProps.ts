import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IProductSearchProps {
  context: WebPartContext;
  intelCategoriesGuid: string;
  involvedNationsGuid: string;
}
