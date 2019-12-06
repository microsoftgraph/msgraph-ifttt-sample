/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { TriggerItem, TriggerResponseData } from "./trigger-response-model";
import { OnenotePage } from "@microsoft/microsoft-graph-types";

/** Creates the response body of the OneNote Page IFTTT Trigger
 *      automatically will set the json payload correctly
 *      with fomatted EventTriggerItem data
 */
export class OnenotePageResponseData implements TriggerResponseData<OnenotePageTriggerItem> {

    data: OnenotePageTriggerItem[];

    constructor (data : OnenotePageTriggerItem[]) {
        this.data = data;
    }

    // Converts from the graph json format to the IFTTT response format
    static fromGraph(data:OnenotePage[], numOfItems:number) : OnenotePageResponseData {

        if (!data || numOfItems === 0) {
            return new OnenotePageResponseData([]);
        }

        // convert the items taking only the properties required.
        let responseData = data.map(
            (dataItem) => {
                return new OnenotePageTriggerItem(dataItem.id, Math.floor(Date.parse(dataItem.createdDateTime)/1000)
                    , dataItem.title
                    , dataItem.parentNotebook.displayName
                    , dataItem.parentSection.displayName
                    , dataItem.contentUrl, dataItem.createdDateTime
                    );
            }
        )

        // Sort the items by timestamp desc (IFTTT requirement)
        responseData.sort((a, b) => b.meta.timestamp - a.meta.timestamp);

        // Filter items to only the requested max amount from IFTTT request
        if (responseData.length > numOfItems)
        {
            responseData = responseData.slice(0, numOfItems);
        }
        console.log(`filtered to ${responseData.length} items`);

        return new OnenotePageResponseData(responseData);
    }
}

/** Creates OneNote Page item data formatted
 *      as required by IFTTT to support the new event triggers
 *      forces the IFTTT requirement of meta data id and timestamp
 *      as well as required sorting and filtering
 */
class OnenotePageTriggerItem extends TriggerItem {
    // tslint:disable: variable-name
    title:string;
    notebook_name:string;
    section_name:string;
    url:string;
    created_at:string;
    // tslint:enable: variable-name

    constructor(id:string, timestamp:number, title:string,
        notebookName:string, sectionName:string, url:string,
        createdAt:string)
    {
        super(id,timestamp);
        this.title = title;
        this.notebook_name = notebookName;
        this.section_name = sectionName;
        this.url = url;
        this.created_at = createdAt;
    }
}
