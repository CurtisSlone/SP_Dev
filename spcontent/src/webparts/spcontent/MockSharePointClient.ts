import { ISPListItem } from "./ISPListItem";

export default class MockSharePointClient {
    private static _listItems: ISPListItem[] = [
        { Id: "1", Title: "First Item"},
        { Id: "2", Title: "Second Item"},
        { Id: "3", Title: "Third Item"},
        { Id: "4", Title: "Fourth Item"},
        { Id: "5", Title: "Fifth Item"},
        { Id: "6", Title: "Sixth Item"},

    ];

    public static get(restUrl: string, options?: any): Promise<ISPListItem[]> {
        return new Promise<ISPListItem[]>((resolve)=>{
            resolve(MockSharePointClient._listItems);
        })
    }
}