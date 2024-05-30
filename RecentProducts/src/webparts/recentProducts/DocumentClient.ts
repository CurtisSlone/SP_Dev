import { SPHttpClient, ISPHttpClientOptions, SPHttpClientResponse} from '@microsoft/sp-http';
import { IDocument } from "./IDocument";
export default class DocumentClient {
    public getDocuments(numDocs: number ): Promise<IDocument[]> {
        return new Promise<IDocument[]>((resolve)=>{
            resolve();
        });
    }
}