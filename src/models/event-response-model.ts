/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { TriggerItem, TriggerResponseData } from "./trigger-response-model";
import { Event} from "@microsoft/microsoft-graph-types";

/** Creates the response body of the Calendar Event IFTTT Trigger
 *      automatically will set the json payload correctly
 *      with formatted EventTriggerItem data
 */
export class EventResponseData implements TriggerResponseData<EventTriggerItem> {

    data: EventTriggerItem[];

    constructor (data : EventTriggerItem[]) {
        this.data = data;
    }

    // Converts from the graph json format to the IFTTT response format
    static fromGraph(data:Event[], numOfItems:number) : EventResponseData {

        if (!data || numOfItems === 0) {
            return new EventResponseData([]);
        }

        // convert the items taking only the properties required.
        let responseData = data.map(
            (dataItem) => {
                return new EventTriggerItem(dataItem.id, Math.floor(Date.parse(dataItem.createdDateTime)/1000),
                    dataItem.organizer.emailAddress.name, dataItem.subject, (new Date(Date.parse(dataItem.start.dateTime))).toLocaleString())
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

        return new EventResponseData(responseData);
    }
}

/** Creates calendar event item data formatted
 *      as required by IFTTT to support the new event triggers
 *      forces the IFTTT requirement of meta data id and timestamp
 *      as well as required sorting and filtering
 */
class EventTriggerItem extends TriggerItem {
    // tslint:disable: variable-name
    organizer_name:string;
    subject:string;
    start_time:string;
    // tslint:enable: variable-name

    constructor(id:string, timestamp:number, organizer:string, subject:string, startTime:string)
    {
        super(id,timestamp);
        this.organizer_name = organizer;
        this.subject = subject;
        this.start_time = startTime;
    }
}

