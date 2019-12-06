/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { TriggerResponseData, TriggerItem } from "./trigger-response-model";
import { Message } from "@microsoft/microsoft-graph-types";

/** Creates the response body of the Message IFTTT Trigger
 *      automatically will set the json payload correctly
 *      with fomatted MessageTriggerItem data
 */
export class MessageResponseData implements TriggerResponseData<MessageTriggerItem>{
    data:MessageTriggerItem[] = [];

    constructor(messages:MessageTriggerItem[])
    {
        this.data = messages;
    }

    // Converts from the graph json format to the IFTTT response format
    static fromGraph(messages:Message[], numOfItems:number):MessageResponseData{
        if (!messages || numOfItems === 0)
        {
            return new MessageResponseData([]);
        }

        // convert the items taking only the properties required.
        let responseData = messages.map(
            msg=>{
                const resMsg = new MessageTriggerItem(msg.id,  Math.floor(Date.parse(msg.receivedDateTime)/1000))
                resMsg.sender = msg.sender.emailAddress.name;
                resMsg.subject = msg.subject;
                return resMsg;
            }
        );

        // Sort the items by timestamp desc (IFTTT requirement)
        responseData.sort((a, b) => b.meta.timestamp - a.meta.timestamp);

        // Filter items to only the requested max amount from IFTTT request
        if (responseData.length > numOfItems)
        {
            responseData = responseData.slice(0,numOfItems);
        }
        console.log(`filtered to ${responseData.length} items`);

        return new MessageResponseData(responseData);
    }
}

/** Creates message event item data formatted
 *      as required by IFTTT to support the new event triggers
 *      forces the IFTTT requirement of meta data id and timestamp
 *      as well as required sorting and filtering
 */
export class MessageTriggerItem extends TriggerItem{
    sender:string;
    subject:string;
}
