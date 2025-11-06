// Interface for a new rating item in the React Emoji Reaction Rating web part
// Matches the structure of the item to be added to the SharePoint list

export interface IRatingNewItem {
    Id?: number;
    Title: string;
    PageID: string; // GUID of the page
    Pagename: string;
    User: string;
    //Comments: string;
    Rating1?: string;
    Rating2?: string;
    Rating3?: string;
    Rating4?: string;
    Rating5?: string;
}