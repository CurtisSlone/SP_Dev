import DocumentClient from "../DocumentClient";
export interface IRecentProductsProps {
  description: string;
  docCount: number;
  documentClient: DocumentClient;
}
