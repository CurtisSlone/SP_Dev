import { IDocument } from "./IDocument";
export default class DocumentClient {
    public getDocuments(numDocs: number ): Promise<IDocument[]> {
        return new Promise<IDocument[]>((resolve)=>{
            resolve();
        });
    }
}