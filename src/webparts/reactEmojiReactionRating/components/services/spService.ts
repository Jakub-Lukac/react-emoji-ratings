import { WebPartContext } from "@microsoft/sp-webpart-base";

import { sp } from "@pnp/sp";
import { Web, IWeb } from "@pnp/sp/webs";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/views";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/security";
import "@pnp/sp/site-users/web";

import { IItemAddResult, IItemUpdateResult, PagedItemCollection } from "@pnp/sp/items";
import { IRatingNewItem } from "../models/IRatingNewItem";

// Class Services
export default class spService {

    constructor(private context: WebPartContext) {
        // Setup Context to PnPjs
        // Initialize PnPjs with the current context
        sp.setup({
            spfxContext: this.context
        });

        // Initialize
        this.onInit();
    }

    // OnInit Function
    private async onInit() {

    }

    /**
     * Get all items from a specified list
     * @param listName - The name of the list to retrieve items from
     * @param centralizedSiteUrl - The URL of the centralized site, if applicable
     * @returns A promise that resolves to an array of items from the specified list
     */
    public async getListItems(
        listName: string,
        centralizedSiteUrl: string
    ) {
        // Check if centralizedSiteUrl is provided, otherwise use the current web
        const targetWeb = centralizedSiteUrl ? Web(centralizedSiteUrl) : sp.web;

        let allItems: any[] = [];

        // Use PnP's getPaged method to retrieve items from the specified list
        // This will handle pagination automatically
        // and return a PagedItemCollection
        // containing the results and additional paging information
        let pagedItems: PagedItemCollection<any[]> = await targetWeb.lists.getByTitle(listName).items.getPaged();

        // Collect items from the first page
        allItems = allItems.concat(pagedItems.results);

        // Continue fetching next pages while available
        while (pagedItems.hasNext) {
            pagedItems = await pagedItems.getNext();
            allItems = allItems.concat(pagedItems.results);
        }

        console.log(allItems);
        return allItems;
    }

    /**
     * Adds a new rating item to the specified list.
     * @param listName - The name of the list to add the item to.
     * @param RatingRequest - The data for the new rating item.
     * @param centralizedSiteUrl - The URL of the centralized site, if applicable.
     * @returns A promise that resolves to the result of the item addition.
     */
    public async addRatingItem(
        listName: string,
        RatingRequest: IRatingNewItem,
        centralizedSiteUrl: string
    ): Promise<IItemAddResult> {
        const targetWeb = Web(centralizedSiteUrl);
        return await targetWeb.lists.getByTitle(listName).items.add(RatingRequest);
    }

    /**
     * Updates an existing rating item in the specified list.
     * @param listName - The name of the list containing the item to update.
     * @param RatingRequest - The updated data for the rating item.
     * @param ExistingID - The ID of the existing item to update.
     * @param centralizedSiteUrl - The URL of the centralized site, if applicable.
     * @returns A promise that resolves to the result of the item update.
     */
    public async updateRatingItem(
        listName: string,
        RatingRequest: IRatingNewItem,
        ExistingID: any,
        centralizedSiteUrl: string
    ): Promise<IItemUpdateResult> {
        const targetWeb = centralizedSiteUrl ? Web(centralizedSiteUrl) : sp.web;
        return await targetWeb.lists.getByTitle(listName).items.getById(ExistingID).update(RatingRequest);
    }

    /**
     * Creates a new list with specified columns in the centralized site.
     * @param listName - The name of the list to create.
     * @param colListColumns - An array of column names to add to the list.
     * @param centralizedSiteUrl - The URL of the centralized site, if applicable.
     * @returns A promise that resolves to a message indicating the result of the operation.
     */
    public async createList(
        listName: string,
        colListColumns: string[],
        centralizedSiteUrl: string
    ): Promise<string> {

        const targetWeb = Web(centralizedSiteUrl);

        const listExist = await this._checkList(listName, targetWeb);

        // If the list does not exist, create it and add the specified columns
        if (!listExist) {
            // Create the list with the specified name
            // and set it to be a custom list (not a document library)
            // The list will be created in the target web instance
            // The list will have a title field by default
            // and we will add the specified columns to it
            const listAddResult = await targetWeb.lists.add(listName);
            const list = listAddResult.list;

            // Log the successful creation of the list
            console.log(`List ${listName} created successfully.`);

            // The default view is a View object that contains the view properties
            // such as the view ID, title, and fields
            // We will use the default view to add the specified columns to it
            // The default view is retrieved using the get() method
            const view = await list.defaultView.get();

            // Add the specified columns to the list
            // and also add them to the default view
            // The fields are added using the addText or addMultilineText methods
            // depending on whether the column is a text or multiline text field
            for (const fieldName of colListColumns) {
                if (fieldName === "Comments") {
                    // The addMultilineText method takes the field name and number of lines as parameters
                    await list.fields.addMultilineText(fieldName, 6);
                } else {
                    // The addText method takes the field name and maximum length as parameters
                    await list.fields.addText(fieldName, 255);
                }
                // We will loop through the colListColumns array and add each column to the list
                // and the default view
                // The fields are added using the fields property of the list object
                await list.views.getById(view.Id).fields.add(fieldName);
            }

            // Break inheritance without copying existing permissions
            await list.breakRoleInheritance(false, false); // copyRoleAssignments=false, clearSubscopes=false
            console.log(`Role inheritance broken for list ${listName}.`);

            // Grant "Contribute" to "Everyone except external users"
            const everyoneExceptExternalUsers = await targetWeb.siteUsers.getByLoginName("c:0-.f|rolemanager|spo-grid-all-users/c5482be3-f911-45e2-a3d8-dff239ab313f")();
            console.log(`Granting Contribute permission to ${everyoneExceptExternalUsers}`);

            await list.roleAssignments.add(everyoneExceptExternalUsers.Id, 1073741827); // 	1073741827 = Contribute
            console.log(`Contribute permission granted to ${everyoneExceptExternalUsers} on list ${listName}.`);

            return "List with required columns created.";
        }

        return "List already exists";
    }

    /**
     * Checks if a list with the specified name exists in the given web instance.
     * @param listName - The name of the list to check.
     * @param webInstance - The web instance to check the list against.
     * @returns A promise that resolves to true if the list exists, false otherwise.
     */
    private async _checkList(listName: string, webInstance: IWeb): Promise<boolean> {
        try {
            await webInstance.lists.getByTitle(listName).get();
            return true;
        } catch {
            return false;
        }
    }

    public async getEmojiMetaData(imageUrl: string, centralizedSiteUrl: string): Promise<{ Title: string }> {
        const targetWeb = centralizedSiteUrl ? Web(centralizedSiteUrl) : sp.web;

        // Extract server-relative URL from the full image URL
        const url = new URL(imageUrl);
        console.log(`Extracted URL: ${url}`);
        const serverRelativeUrl = decodeURIComponent(url.pathname);
        console.log(`Server-relative URL: ${serverRelativeUrl}`);
        
        try {
            const file = await targetWeb.getFileByServerRelativePath(serverRelativeUrl).select("Title").get();
            console.log(`Fetched metadata for ${imageUrl}:`, file);
            return { Title: file.Title };
        } catch (error) {
            console.error(`Error fetching metadata for ${imageUrl}:`, error);
            return { Title: "" };
        }
    }
}
