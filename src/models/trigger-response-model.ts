/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/** Interface for configuring the response body of any IFTTT Trigger
 *      automatically will set the json payload correctly
 *      with formatted TriggerItem data
 */
export interface TriggerResponseData<T extends TriggerItem> {

    data : T[];
}

/** Creates item data formatted for the response body of any IFTTT Trigger
 *      forces the IFTTT requirement of meta data id and timestamp
 */
export class TriggerItem implements Item {
    /**
     * adding public properties here in a derived class
     *  will automatically add these items to the json payload.
     */

    constructor (id:string, timestamp: number) {
        this.meta = new Meta(id, timestamp);
    }
    meta: Meta;
}

export interface Item {

    meta: Meta;
}

export class Meta {

    constructor (id:string, timestamp:number) {
        this.id = id;
        this.timestamp = timestamp;
    }
    id: string;
    timestamp: number
}

